from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


import inspect
from openpyxl.compat import unicode, basestring, zip
from openpyxl.descriptors import Descriptor
from openpyxl.descriptors.serialisable import Serialisable


class HashableObject(Serialisable):
    """Define how to hash property classes."""
    __fields__ = ()
    __base__ = False
    _key = None

    @property
    def __defaults__(self):
        spec = inspect.getargspec(self.__class__.__init__)
        return dict(zip(spec.args[1:], spec.defaults))

    def copy(self, **kwargs):
        current = dict([(x, getattr(self, x)) for x in self.__fields__])
        current.update(kwargs)
        return self.__class__(**current)

    @property
    def key(self):
        """Use a tuple of fields as the basis for a key"""
        if self._key is None:
            fields = []
            for attr in self.__fields__:
                val = getattr(self, attr)
                if isinstance(val, list):
                    val = tuple(val)
                fields.append(val)
            self._key = hash(tuple(fields))
        return self._key

    def __hash__(self):
        return self.key

    def __eq__(self, other):
        if other.__class__ == self.__class__:
            return self.key == other.key
        return self.key == other

    def __ne__(self, other):
        return not self == other

    def __add__(self, other):
        vals = {}
        for attr in self.__fields__:
            vals[attr] = getattr(self, attr) or getattr(other, attr)
        return self.__class__(**vals)

    def __sub__(self, other):
        vals = {}
        if (self is other) or (self == other):
            return
        for attr in self.__fields__:
            if not getattr(other, attr) and getattr(self, attr):
                vals[attr] = getattr(self, attr)
        return self.__class__(**vals)
