import unittest

import popup_positioning


class FakeParent:
    def update_idletasks(self):
        pass

    def winfo_rootx(self):
        return -1920

    def winfo_rooty(self):
        return 90

    def winfo_width(self):
        return 1200

    def winfo_height(self):
        return 800


class FakeWindow:
    def __init__(self):
        self.geometry_value = None
        self.transient_parent = None

    def update_idletasks(self):
        pass

    def winfo_width(self):
        return 600

    def winfo_height(self):
        return 400

    def winfo_reqwidth(self):
        return 600

    def winfo_reqheight(self):
        return 400

    def geometry(self, value):
        self.geometry_value = value

    def transient(self, parent):
        self.transient_parent = parent


class PopupPositioningTests(unittest.TestCase):
    def test_popup_is_centered_on_a_left_hand_secondary_monitor(self):
        parent = FakeParent()
        window = FakeWindow()

        popup_positioning.place_over_parent(window, parent)

        self.assertEqual("-1620+290", window.geometry_value)
        self.assertIs(parent, window.transient_parent)
