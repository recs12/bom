"""
Api solidedge
=======================

"""

import clr

clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")
clr.AddReference("System")

from System import Console
import System.Runtime.InteropServices as SRI


def raw_input(message):
    Console.WriteLine(message)
    return Console.ReadLine()
class Api:
    def __init__(self):
        # Connect to a running instance of Solid Edge
        self.api = SRI.Marshal.GetActiveObject("SolidEdge.Application")

    def check_valid_version(self, *valid_version):
        # validate solidedge version - 'Solid Edge ST7'
        print("version solidedge: %s" % self.api.Value)
        assert self.api.Value in valid_version, "Unvalid version of solidedge"

    def active_document(self):
        return self.api.ActiveDocument

    @property
    def name(cls):
        return cls.part.Name

    def delete(cls):
        return cls.part.Delete()
