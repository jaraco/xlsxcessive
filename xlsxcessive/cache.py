class CacheDecorator:
    def __init__(self):
        self.cache = {}
        self.func = None

    def cached_func(self, *args):
        if args not in self.cache:
            self.cache[args] = self.func(*args)
        return self.cache[args]

    def __call__(self, func):
        self.func = func
        return self.cached_func
