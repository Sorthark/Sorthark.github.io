from html.parser import HTMLParser
from pathlib import Path


class Checker(HTMLParser):
    void = {"meta", "link", "img", "br", "hr", "input", "area", "base", "col", "embed", "source", "track", "wbr"}

    def __init__(self):
        super().__init__()
        self.stack: list[str] = []
        self.errors: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag not in self.void:
            self.stack.append(tag)

    def handle_endtag(self, tag):
        if tag in self.void:
            return
        if not self.stack:
            self.errors.append(f"unexpected </{tag}>")
            return
        if self.stack[-1] != tag:
            self.errors.append(f"</{tag}> but expected </{self.stack[-1]}>")
        else:
            self.stack.pop()


html = Path(__file__).resolve().parents[1].joinpath("index.html").read_text(encoding="utf-8")
checker = Checker()
checker.feed(html)
print("articles", html.count('class="product-card'))
print("unclosed", checker.stack)
print("errors", checker.errors[:10], "count", len(checker.errors))
