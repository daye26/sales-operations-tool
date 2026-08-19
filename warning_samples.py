"""Bound warning detail retention while preserving reported totals."""


class WarningSamples:
    def __init__(self, limit=20):
        self._limit = limit
        self._count = 0
        self._samples = []

    def append(self, value):
        self._count += 1
        if len(self._samples) < self._limit:
            self._samples.append(value)

    def __bool__(self):
        return self._count > 0

    def __len__(self):
        return self._count

    def __getitem__(self, index):
        return self._samples[index]
