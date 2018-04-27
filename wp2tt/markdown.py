#!/usr/bin/env python3
import logging  # noqa: F401
import collections
import contextlib
import mistune
from lxml import etree
from wp2tt.input import IDocumentFootnote
from wp2tt.input import IDocumentInput
from wp2tt.input import IDocumentParagraph
from wp2tt.input import IDocumentSpan
from wp2tt.styles import DocumentProperties


class Elements(list):
    def __iadd__(self, other):
        self.append(other)
        return self


class MarkdownUnRenderer(object):
    def __init__(self, **kwargs):
        self.options = kwargs
        self.wpids = collections.defaultdict(set)
        self.root = etree.Element('document')

    def placeholder(self):
        return Elements()

    def text(self, text):
        return etree.Element('span', text=text)

    def paragraph(self, elements):
        node = etree.SubElement(self.root, 'paragraph')
        for span in elements:
            node.append(span)
        return node

    def emphasis(self, nodes):
        return self._typed_span(nodes, 'emphasis')

    def double_emphasis(self, nodes):
        return self._typed_span(nodes, 'double_emphasis')

    def _typed_span(self, nodes, wpid):
        node = nodes[-1]
        node.attrib['wpid'] = wpid
        self.wpids['character'].add(wpid)
        return node

    def block_code(self, code, language=None):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_code'))

    def block_quote(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_quote'))

    def block_html(self, html):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'block_html'))

    def header(self, text, level, raw=None):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'header'))

    def hrule(self, ):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'hrule'))

    def list(self, body, ordered=True):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'list'))

    def list_item(self, text):
        raise NotImplementedError('%s.%s()' % (type(self).__name__, 'list_item'))

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
        self._properties = DocumentProperties()

    def _read_markdown(self, path):
        self._markdown = MarkdownUnRenderer()
        parse = mistune.Markdown(renderer=self._markdown)
        with open(path, 'r') as md:
            parse(md.read())
        logging.debug(etree.tostring(self._markdown.root))

    @property
    def properties(self):
        return self._properties

    def styles_defined(self):
        """Yield a Style object kwargs for every style defined in the document."""
        for realm, wpid in self.styles_in_use():
            yield {'realm': realm, 'internal_name': wpid, 'wpid': wpid}

    def styles_in_use(self):
        """Yield a pair (realm, wpid) for every style used in the document."""
        for realm, realm_wpids in self._markdown.wpids.items():
            for wpid in realm_wpids:
                yield (realm, wpid)

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each body paragraph."""
        for p in self._markdown.root:
            yield MarkdownParagraph(p)


class MarkdownParagraph(IDocumentParagraph):
    """A Paragraph inside a document."""

    def __init__(self, node):
        self.node = node

    def style_wpid(self):
        """Returns the wpid for this paragraph's style."""
        return self.node.get('wpid')

    def text(self):
        """Yields strings of plain text."""
        raise NotImplementedError()

    def spans(self):
        """Yield a MarkdownSpan per text span."""
        for span in self.node:
            yield MarkdownSpan(span)


class MarkdownSpan(IDocumentSpan):
    """A span of characters inside a document."""

    def __init__(self, node):
        self.node = node

    def style_wpid(self):
        """Returns the wpid for this span's style."""
        return self.node.get('wpid')

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
        yield self.node.get('text')


class MarkdownFootnote(IDocumentFootnote):
    """A footnote."""

    def paragraphs(self):
        """Yields a MarkdownParagraph object for each footnote paragraph."""
        raise NotImplementedError()
