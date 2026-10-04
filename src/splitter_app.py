from __future__ import annotations

from splitter_mixin_1 import SplitterMixin1
from splitter_mixin_2 import SplitterMixin2
from splitter_mixin_3 import SplitterMixin3

class SplitterApp(SplitterMixin1, SplitterMixin2, SplitterMixin3):
    """界面实现拆在 mixin 中，行为与原来的单个类相同。"""
