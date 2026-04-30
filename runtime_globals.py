import builtins
import copy


class RuntimeGlobal:
    def __init__(self, name, default=None):
        self.name = name
        self.default = default

    def _value(self):
        return getattr(builtins, self.name, self.default)

    def __getitem__(self, key):
        return self._value()[key]

    def __contains__(self, item):
        return item in self._value()

    def __iter__(self):
        return iter(self._value())

    def __len__(self):
        return len(self._value())

    def __bool__(self):
        return bool(self._value())

    def __str__(self):
        return str(self._value())

    def __repr__(self):
        return repr(self._value())

    def __getattr__(self, name):
        return getattr(self._value(), name)

    def __deepcopy__(self, memo):
        return copy.deepcopy(self._value(), memo)

    def get(self, key, default=None):
        return self._value().get(key, default)


configContent = RuntimeGlobal('configContent', {})
staff_dict = RuntimeGlobal('staff_dict', {})
today = RuntimeGlobal('today', '')
desktopUrl = RuntimeGlobal('desktopUrl', '')
myWin = RuntimeGlobal('myWin')
myTable = RuntimeGlobal('myTable')
