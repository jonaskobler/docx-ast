from docx_ast.ast_to_docx import ASTDocxSerializer
from docx_ast.ast_to_str import ASTStringSerializer
from docx_ast.docx_to_ast import DocxASTExtractor
from docx_ast.nodes import (
    Comment,
    Document,
    Footnote,
    Node,
    Paragraph,
    Text,
    link_parents,
    normalize_ws,
    walk,
)
from docx_ast.tracked_changes import alter_paragraph_tracked, alter_text_tracked

__version__ = "0.1.0"

__all__ = [
    "ASTDocxSerializer",
    "ASTStringSerializer",
    "DocxASTExtractor",
    "Comment",
    "Document",
    "Footnote",
    "Node",
    "Paragraph",
    "Text",
    "link_parents",
    "normalize_ws",
    "walk",
    "alter_paragraph_tracked",
    "alter_text_tracked",
]
