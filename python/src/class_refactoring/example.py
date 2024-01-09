"""
This module contains simple classes that demonstrate how to change from an method's original type (instance, class or static) to another.


"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from datetime import datetime
from enum import Enum, verify, UNIQUE
from typing import List

import numpy as np
import numpy.typing as npt


@verify(UNIQUE)
class OptionType(Enum):
    """
    Enumeration covering some common option types
    """
    AMERICAN = 1
    EUROPEAN = 2
    ASIAN = 3
    BERMUDAN = 4


class RandomUnderlying:
    """
    A trivial representation of an option underlying
    """
    def __init__(self, spot: float, vol: float, maturity: datetime):
        self.spot = spot
        self.vol = vol
        self.maturity = maturity

    @property
    def current_spot(self):
        """
        Underlying spot price
        """
        return self.spot

    @property
    def volatility(self):
        """
        Underlying volatility
        """
        return self.vol

@dataclass
class SomeOption:
    """
    A bare bone representation of an option
    """
    option_type: OptionType
    strike: float
    expiry_date: datetime


class OptionPricingEngine(ABC):
    """
    An abstract base class for a pricing engine
    """
    @classmethod
    @abstractmethod
    def check_options(cls, options: List[SomeOption]):
        """
        Method checking a list of options to determine whether the engine can price them
        """

    @abstractmethod
    def price_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]) -> npt.NDArray[np.float64]:
        """
        Method returning prices for a list of options
        """


class EuropeanBlackScholesPricingEngine(OptionPricingEngine):
    """
    A pricer implementation that only handles European options
    """
    accepted_type = OptionType.EUROPEAN

    @classmethod
    def check_options(cls, options: List[SomeOption]):
        check = [opt for opt in options if opt.option_type == cls.accepted_type]
        if len(check) != len(options):
            raise RuntimeError("Options list contains invalid types")

    def price_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        normal_dist = np.random.default_rng().normal

        spot = underlying.current_spot, vol = underlying.volatility
        ttm = np.array([(underlying.maturity - opt.expiry_date).days for opt in options])
        scaled_vol = vol * np.sqrt(ttm)
        strikes = np.array([opt.strike for opt in options])

        d1 = np.log(np.divide(spot, strikes)) + np.multiply((risk_free_rate + 0.5 * vol**2), ttm)
        d1 = np.divide(d1, scaled_vol)

        d2 = d1 - scaled_vol
        return normal_dist(d1) * spot - normal_dist(d2) * strikes * np.exp(-risk_free_rate * ttm)


class FirstRefactoredPricingEngine(OptionPricingEngine):
    """
    A refactoring of the EuropeanBlackScholesPricingEngine class that attempts to cover more options types, and redefine check_options as an instance method
    """
    def __init__(self):
        self.accepted_types = [OptionType.EUROPEAN, OptionType.AMERICAN]

    def check_options(self, options: List[SomeOption]):
        check = [opt for opt in options if opt.option_type in self.accepted_types]
        if len(check) != len(options):
            invalid_types = [opt.option_type for opt in options if opt.option_type not in self.accepted_types]
            raise RuntimeError(f"Options list contains invalid types: {invalid_types}")

    def price_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        european = [opt for opt in options if opt.option_type == OptionType.EUROPEAN]
        american = [opt for opt in options if opt.option_type == OptionType.AMERICAN]

        european_prices = self._price_european_options(underlying, risk_free_rate, european) if european else []
        american_prices = self._price_american_options(underlying, risk_free_rate, american) if american else []

        return european_prices + american_prices

    def _price_american_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        """
        """
        raise NotImplementedError("American options pricing not yet implemented")

    def _price_european_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        """
        """
        raise NotImplementedError("European options pricing not yet implemented")


class SecondRefactoredPricingEngine(OptionPricingEngine):
    """
    A second refactoring of EuropeanBlackScholesPricingEngine with class and instance check_options methods
    """
    accepted_types = [OptionType.EUROPEAN, OptionType.AMERICAN]

    def __init__(self, accepted_types: List[OptionType] = None):
        self._accepted_types = accepted_types or type(self).accepted_types
        # Using self.check_options = self._instance_check_options will usually work, unless a linter check complains about function hiding
        setattr(self, 'check_options', self._instance_check_options)

    @classmethod
    def check_options(cls, options: List[SomeOption]):  # pylint: disable=missing-function-docstring
        check = [opt for opt in options if opt.option_type in cls.accepted_types]
        if len(check) != len(options):
            invalid_types = [str(opt.option_type) for opt in options if opt.option_type not in cls.accepted_types]
            raise RuntimeError(f"Options list contains invalid types: {invalid_types}")

    def _instance_check_options(self, options: List[SomeOption]):
        check = [opt for opt in options if opt.option_type in self._accepted_types]
        if len(check) != len(options):
            invalid_types = [str(opt.option_type) for opt in options if opt.option_type not in self._accepted_types]
            raise RuntimeError(f"Options list contains invalid types: {invalid_types}")

    def price_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):  # pylint: disable=missing-function-docstring
        european = [opt for opt in options if opt.option_type == OptionType.EUROPEAN]
        american = [opt for opt in options if opt.option_type == OptionType.AMERICAN]

        european_prices = self._price_european_options(underlying, risk_free_rate, european) if european else []
        american_prices = self._price_american_options(underlying, risk_free_rate, american) if american else []

        return european_prices + american_prices

    def _price_american_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        """
        """
        raise NotImplementedError("American options pricing not yet implemented")

    def _price_european_options(self, underlying: RandomUnderlying, risk_free_rate: float, options: List[SomeOption]):
        """
        """
        raise NotImplementedError("European options pricing not yet implemented")
