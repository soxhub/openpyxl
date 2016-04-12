from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl.compat import OrderedDict
from .rule import Rule


class ConditionalFormatting(object):
    """Conditional formatting rules."""

    def __init__(self):
        self.cf_rules = OrderedDict()
        self.max_priority = 0

    def add(self, range_string, cfRule):
        """Add a rule such as ColorScaleRule, FormulaRule or CellIsRule

         The priority will be added automatically.
        """
        if not isinstance(cfRule, Rule):
            raise ValueError("Only instances of openpyxl.formatting.rule.Rule may be added")
        rule = cfRule
        self.max_priority += 1
        rule.priority = self.max_priority

        self.cf_rules.setdefault(range_string, []).append(rule)
