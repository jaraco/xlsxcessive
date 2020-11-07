from xlsxcessive.style import Stylesheet


class TestStylesheetNumbering:
    def setup(self):
        self.styles = Stylesheet(None)

    def test_can_add_custom_number_format(self):
        self.styles.add_custom_number_format("#.0000")
        assert "#.0000" in self.styles.custom_numbers

    def test_adding_custom_styles_returns_an_int(self):
        actual = type(self.styles.add_custom_number_format("#.0000"))
        expected = int
        assert actual == expected

    def test_adding_custom_styles_increments_the_int(self):
        first = self.styles.add_custom_number_format("#.0000")
        second = self.styles.add_custom_number_format("#.)0000")
        assert second > first

    def test_adding_the_same_style_twice_does_not_increment(self):
        first = self.styles.add_custom_number_format("#.0000")
        second = self.styles.add_custom_number_format("#.0000")
        assert first == second, (first, second)


class TestFormatNumbering:
    def setup(self):
        self.styles = Stylesheet(None)
        self.format = self.styles.new_format()

    def test_using_a_builtin_format_does_not_add_a_custom_format(self):
        self.format.number_format('0.00')
        assert not self.styles.custom_numbers

    def test_using_a_custom_format_does_add_a_custom_format(self):
        self.format.number_format('0.00000')
        assert self.styles.custom_numbers

    def test_format_codes_are_escaped(self):
        self.format.number_format('"x"0.00')
        assert '&quot;x&quot;0.00' in self.styles.custom_numbers
