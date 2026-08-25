"""Position custom Tk dialogs relative to the Sales Operations main window."""


def place_over_parent(window, parent):
    """Center a dialog on its parent, including on secondary monitors."""
    parent.update_idletasks()
    window.update_idletasks()
    width = max(window.winfo_width(), window.winfo_reqwidth())
    height = max(window.winfo_height(), window.winfo_reqheight())
    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    x_offset = f"+{x}" if x >= 0 else str(x)
    y_offset = f"+{y}" if y >= 0 else str(y)
    window.geometry(f"{x_offset}{y_offset}")
    window.transient(parent)
