class StructureCorruptedError(Exception):
    """Exception raised for non-conform XML stuff"""

    def __init__(self):
        super().__init__(
            "We encoutered an impossible relation in the XML, the structure is corrupted."
        )
