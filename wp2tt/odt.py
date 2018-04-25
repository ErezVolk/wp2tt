#!/usr/bin/env python
import contextlib
import logging
import zipfile

import lxml.etree

from wp2tt.styles import DocumentProperties


class OoXml(object):
    """Basic helper class for the OpenOffice XML format."""
    _NS = {
        'office': 'urn:oasis:names:tc:opendocument:xmlns:office:1.0',
        'style': 'urn:oasis:names:tc:opendocument:xmlns:style:1.0',
        'text': 'urn:oasis:names:tc:opendocument:xmlns:text:1.0',
    }

    _FAMILY_TO_REALM = {
        'paragraph': 'paragraph',
        'text': 'character',
    }

    def _xpath(self, node, expr):
        return node.xpath(expr, namespaces=self._NS)

    def _ootag(self, tag):
        ns, tag = tag.split(':', 1)
        return '{%s}%s' % (self._NS[ns], tag)

    def _ooget(self, node, tag):
        return node.get(self._ootag(tag))


class OdtInput(contextlib.ExitStack, OoXml):
    """A .docx reader."""
    def __init__(self, path):
        super().__init__()
        self._zip = self.enter_context(zipfile.ZipFile(path))
        self._content = self._load_xml('content.xml')
        self._initialize_properties()

    def _initialize_properties(self):
        self._properties = DocumentProperties(
            has_rtl=self._has_node('//style:paragraph-properties[@style:writing-mode="rl-tb"]')
        )

    def _has_node(self, ootag):
        for node in self._xpath(self._content, ootag):
            return True
        return False

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        for s in self._xpath(self._load_xml('styles.xml'), '//office:styles/style:style'):
            yield self._style_kwargs(s)
        for s in self._xpath(self._content, '//office:automatic-styles/style:style'):
            yield self._style_kwargs(s, automatic=True)

    def _style_kwargs(self, node, **extras):
        name = self._ooget(node, 'style:name')
        style_kwargs = dict(extras) if extras else {}
        style_kwargs.update({
            'realm': self._FAMILY_TO_REALM[self._ooget(node, 'style:family')],
            'internal_name': self._ooget(node, 'style:display-name') or name,
            'wpid': name,
            'parent_wpid': self._ooget(node, 'style:parent-style-name'),
            'next_wpid': self._ooget(node, 'style:next-style-name'),
        })
        return style_kwargs

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for realm, tag in (('paragraph', 'text:p'), ('character', 'text:span')):
            for sn in self._xpath(self._content, '//%s' % tag):
                wpid = self._ooget(sn, 'text:style-name')
                yield (realm, wpid)

    def paragraphs(self):
        """Yields a OdtParagraph object for each body paragraph."""
        for p in self._xpath(self._content, '//office:body/office:text/text:p'):
            yield OdtParagraph(self, p)

    def _load_xml(self, path_in_zip):
        """Parse an XML file inside the zipped doc, return root node."""
        try:
            with self._zip.open(path_in_zip) as fo:
                return lxml.etree.parse(fo).getroot()
        except KeyError:
            return None


class OdtNode(OoXml):
    """Base helper class for object which represent a node in a docx."""
    def __init__(self, doc, node):
        self.doc = doc
        self.node = node

    def _node_ooget(self, tag):
        return self._ooget(self.node, tag)

    def _node_xpath(self, expr):
        return self.node.xpath(expr, namespaces=self._NS)


class OdtParagraph(OdtNode):
    """A Paragraph inside a .docx."""
    def style_wpid(self):
        return self._node_ooget('text:style-name')

    def text(self):
        """Yields strings of plain text."""
        for t in self.itertext():
            yield t

    def spans(self):
        """Yield OdtSpan per text span."""
        for event, n in lxml.etree.iterwalk(self.node, events=('start', 'end')):
            if event == 'start':
                if n.tag == self._ootag('text:tab'):
                    yield OdtTabSpan(self.doc, n)
                elif n.tag == self._ootag('text:span'):
                    yield OdtSpanSpan(self.doc, n)
                elif n.tag == self._ootag('text:p'):
                    yield OdtHeadSpan(self.doc, n)
                else:
                    logging.debug('Not sure what to do with a <%s> %r', n.tag, n.text[:8])
                    yield OdtHeadSpan(self.doc, n)
            else:
                yield OdtTailSpan(self.doc, n)


class OdtSpanBase(OdtNode):
    def style_wpid(self):
        return None

    def footnotes(self):
        if False:
            yield

    def comments(self):
        if False:
            yield

    def text(self):
        if False:
            yield


class OdtHeadSpan(OdtSpanBase):
    def text(self):
        if self.node.text:
            yield self.node.text


class OdtTabSpan(OdtSpanBase):
    def text(self):
        yield '\t'


class OdtSpanSpan(OdtSpanBase):
    def style_wpid(self):
        return self._node_ooget('text:style-name')

    def footnotes(self):
        for fnr in self._node_xpath('text:note[@text:node-class="footnote"]'):
            yield OdtFootnote(self.doc, fnr)

    def comments(self):
        if False:
            yield

    def text(self):
        if self.node.text:
            yield self.node.text


class OdtTailSpan(OdtSpanBase):
    def text(self):
        if self.node.tail:
            yield self.node.tail


class OdtFootnote(OdtNode):
    def __init__(self, doc, node):
        super().__init__(doc, node)

    def paragraphs(self):
        for p in self._node_xpath('text:note-body/text-p'):
            yield OdtParagraph(self.doc, p)


