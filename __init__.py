# -*- coding: utf-8 -*-

from .halo import Halo

def classFactory(iface):
    """Load Halo class from file halo.py"""
    return Halo(iface)