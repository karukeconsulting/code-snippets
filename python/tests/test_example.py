from typing import List
from datetime import datetime
from class_refactoring.example import SomeOption, OptionType, OptionPricingEngine, EuropeanBlackScholesPricingEngine, \
    FirstRefactoredPricingEngine, SecondRefactoredPricingEngine

import re
import pytest


def test_class_method_refactoring():
    options = [
        SomeOption(OptionType.AMERICAN, 43.5, datetime(2024, 5, 3)),
        SomeOption(OptionType.BERMUDAN, 31.7, datetime(2024, 5, 3)),
        SomeOption(OptionType.EUROPEAN, 10.3, datetime(2024, 5, 3))
    ]

    UsedCls = EuropeanBlackScholesPricingEngine
    with pytest.raises(RuntimeError, match="Options list contains invalid types"):
        UsedCls.check_options(options)

    with pytest.raises(RuntimeError, match="Options list contains invalid types"):
        UsedCls().check_options(options)

    UsedCls = FirstRefactoredPricingEngine
    with pytest.raises(TypeError, match=re.escape(f"{UsedCls.__name__}.check_options() missing 1 required positional argument: 'options'")):
        UsedCls.check_options(options)

    with pytest.raises(RuntimeError, match="Options list contains invalid types"):
        UsedCls().check_options(options)

    UsedCls = SecondRefactoredPricingEngine
    with pytest.raises(RuntimeError, match=re.escape("Options list contains invalid types: ['OptionType.BERMUDAN']")):
        UsedCls.check_options(options)

    with pytest.raises(RuntimeError, match="Options list contains invalid types"):
        UsedCls().check_options(options)

    assert UsedCls([OptionType.AMERICAN, OptionType.BERMUDAN, OptionType.EUROPEAN]).check_options(options) is None
