#!/usr/bin/env python3
import logging  # noqa: F401
import contextlib
import mistune
from lxml import etree
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.styles import DocumentProperties


class Element(object):
    def __init__(self, *args, **kwargs):
        if not args:
            args = ['nop']
        self.e = etree.Element(*args, **kwargs)

    def __iter__(self):
        return self.e.__iter__()

    def __getitem__(self, index):
        return self.e.__getitem__(index)

    def __iadd__(self, other):
        try:
            self.e.append(other.e)
        except AttributeError:
            self.e.append(other)
        return self


class MarkdownUnRenderer(object):
    def __init__(self, **kwargs):
        self.options = kwargs

    def placeholder(self):
        return ''

    def header(self, text, level, raw=None):
        return '<p wpid="header">%s</p>' % text

    def text(self, text):
        return '<s>%s</s>' % text

    def paragraph(self, text):
        return '<p>%s</p>' % text

    def emphasis(self, text):
        return '<s wpid="emphasis">%s</s>' % text

    def double_emphasis(self, text):
        return '<s wpid="double emphasis">%s</s>' % text

    def list_item(self, text):
        return '<li>%s</li>' % text

    def list(self, text, ordered=True):
        return text

    def block_code(self, code, language=None):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_code'))

    def block_quote(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_quote'))

    def block_html(self, html):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_html'))

    def hrule(self, ):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'hrule'))

    def table(self, header, body):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'table'))

    def table_row(self, content):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'table_row'))

    def table_cell(self, content, **flags):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'table_cell'))

    def autolink(self, link, is_email=False):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'autolink'))

    def codespan(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'codespan'))

    def image(self, src, title, alt_text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'image'))

    def linebreak(self, ):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'linebreak'))

    def newline(self, ):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'newline'))

    def link(self, link, title, content):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'link'))

    def strikethrough(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'strikethrough'))

    def inline_html(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'inline_html'))

    def footnote_ref(self, key, index):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'footnote_ref'))

    def footnote_item(self, key, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'footnote_item'))

    def footnotes(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'footnotes'))


class MarkdownInput(contextlib.ExitStack, IDocumentInput):
    """A Markdown reader."""
    def __init__(self, path):
        super().__init__()
        self._read_markdown(path)
        self._properties = DocumentProperties(has_rtl=False)

    def _read_markdown(self, path):
        renderer = MarkdownUnRenderer()
        parse = mistune.Markdown(renderer=renderer)
        with open(path, 'r') as md:
            xml = parse(md.read())
        self._root = etree.fromstring('<document>%s</document>' % xml)
        print(etree.tostring(self._root, pretty_print=True, encoding='utf-8', xml_declaration=True).decode('utf-8'))

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        for realm, wpid in self.styles_in_use():
            yield {'realm': realm, 'internal_name': wpid, 'wpid': wpid}

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for node in self._root.xpath('//p[@wpid]'):
            yield 'paragraph', node.get('wpid')
        for node in self._root.xpath('//s[@wpid]'):
            yield 'character', node.get('wpid')
        for node in self._root.xpath('//li'):
            yield 'paragraph', 'list item'
            break

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each body paragraph."""
        for p in self._root:
            if p.tag in ('li', 'p'):
                yield MarkdownParagraph(p)


class MarkdownParagraph(IDocumentParagraph):
    """A Paragraph inside a document."""

    def __init__(self, node):
        self.node = node
        if node.tag == 'p':
            self.wpid = node.get('wpid')
        if node.tag == 'li':
            self.wpid = 'list item'
            if len(node) and node[0].tag == 'p':
                self.node = node[0]

    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        return self.wpid

    def text(self):
        """Yields strings of plain text."""
        for span in self.spans():
            for text in span.text():
                yield span.text()

    def spans(self):
        """Yield a MarkdownSpan per text span."""
        for span in self.node.xpath('s'):
            yield MarkdownSpan(span)


class MarkdownSpan(IDocumentSpan):
    """A span of characters inside a document."""

    def __init__(self, node):
        self.wpid = node.get('wpid')
        while len(node) and node[0].tag == 's':
            node = node[0]
        self.node = node

    def style_wpid(self):
        """Returns the wpid for this span's style."""
        return self.wpid

    def footnotes(self):
        """Yields a MarkdownFootnote object for each footnote in this span."""
        if False:
            yield

    def comments(self):
        """Yields a MarkdownComment object for each comment in this span."""
        if False:
            yield

    def text(self):
        """Yields strings of plain text."""
        yield self.node.text or ''


class MarkdownFootnote(IDocumentFootnote):
    """A footnote."""

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each footnote paragraph."""
        raise NotImplementedError()
