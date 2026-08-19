"""Shared source precedence for operational vehicle ports."""


def value_text(value):
    if value is None:
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()


def resolve_port(
    port_stock_port=None,
    newport_port=None,
    not_allocated_port=None,
    vehicle_tracking_port=None,
):
    """Use the agreed operational port precedence for a VIN."""
    for port in (
        port_stock_port,
        newport_port,
        not_allocated_port,
        vehicle_tracking_port,
    ):
        normalized = value_text(port)
        if normalized:
            return normalized
    return ""
