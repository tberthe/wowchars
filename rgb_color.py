# -*- coding: utf-8 -*-


class RGBColor:
    """Helper class to easily deal with colors"""
    def __init__(self, red, green, blue):
        """Constructor

        Args:
            red (int): red integer value (from 0 to 255)
            green (int): green integer value (from 0 to 255)
            blue (int): blue integer value (from 0 to 255)
        """
        self.red = red
        self.green = green
        self.blue = blue

    @classmethod
    def from_float_rgb(fgbc, fred, fgreen, fblue):
        """Factory to create RGBColor from float values

        Args:
            fgbc (int): red float value (from 0 to 1)
            fgreen (int): green float value (from 0 to 1)
            fblue (int): blue float value (from 0 to 1)

        Returns:
            a RGBColor object
        """
        r = int(int(round(fred * 255)))
        g = int(int(round(fgreen * 255)))
        b = int(int(round(fblue * 255)))
        return fgbc(r, g, b)

    @classmethod
    def from_float_rgb_dict(fgbc, d):
        """Factory to create RGBColor from float values

        Args:
            d (dict): dictionnary containing the red/green/blue float values

        Returns:
            a RGBColor object
        """
        r = d["red"] if "red" in d else 0.0
        g = d["green"] if "green" in d else 0.0
        b = d["blue"] if "blue" in d else 0.0
        return fgbc.from_float_rgb(r, g, b)

    @classmethod
    def from_hex(fgbc, hex_code):
        """Factory to create RGBColor from the hexadecimal value

        Args:
            hex_code (str): hexadecimal string of the color, with or without hash. (ex: "#FF0000")

        Returns:
            a RGBColor object
        """
        h = hex_code.lstrip('#')
        rgb = tuple(int(h[i: i + 2], 16) for i in (0, 2, 4))
        return fgbc(rgb[0], rgb[1], rgb[2])

    def to_hex(self):
        """Returns: color as a hexadecimal string (ex: "#FF0000")"""
        return '#%02X%02X%02X' % (self.red, self.green, self.blue)

    def to_float_rgb_dict(self):
        """Returns: color as a dictionnary with float values"""
        return {"red": self.red / 255.0,
                "green": self.green / 255.0,
                "blue": self.blue / 255.0}

    def to_rgb_dict(self):
        """Returns: color as a dictionnary with int values"""
        return {"red": self.red,
                "green": self.green,
                "blue": self.blue}

    def __eq__(self, other):
        """Equality operator"""
        return (self.red == other.red) and (self.green == other.green) and (self.blue == other.blue)
