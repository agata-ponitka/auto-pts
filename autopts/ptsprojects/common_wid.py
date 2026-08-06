#
# auto-pts - The Bluetooth PTS Automation Framework
#
# Copyright (c) 2026, Codecoup.
#
# This program is free software; you can redistribute it and/or modify it
# under the terms and conditions of the GNU General Public License,
# version 2, as published by the Free Software Foundation.
#
# This program is distributed in the hope it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
# FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for
# more details.
#

import importlib.util
import logging
from enum import Enum

from autopts.ptsprojects.stack import get_stack

log = logging.debug


class Backend(Enum):
    ZEPHYR = "zephyr"
    MYNEWT = "mynewt"
    BLUEZ = "bluez"


class Profile(Enum):
    AICS = "aics"
    ASCS = "ascs"
    BAP = "bap"
    BASS = "bass"
    CAP = "cap"
    CAS = "cas"
    CCP = "ccp"
    CSIP = "csip"
    CSIS = "csis"
    DIS = "dis"
    GAP = "gap"
    GATT = "gatt"
    GATT_CLIENT = "gatt_client"
    GATTC = "gattc"
    GMCS = "gmcs"
    GTBS = "gtbs"
    HAP = "hap"
    HAS = "has"
    IAS = "ias"
    L2CAP = "l2cap"
    MCP = "mcp"
    MESH = "mesh"
    MICP = "micp"
    MICS = "mics"
    MMDL = "mmdl"
    OTS = "ots"
    PACS = "pacs"
    PBP = "pbp"
    RFCOMM = "rfcomm"
    SDP = "sdp"
    SM = "sm"
    TBS = "tbs"
    TMAP = "tmap"
    VCP = "vcp"
    VCS = "vcs"
    VOCS = "vocs"
    # GENERATOR append 1


def _backend_wid_module(backend: Backend, profile: Profile):
    """Return the module name for <backend>/<profile>_wid.py, or None if absent."""
    name = f"autopts.ptsprojects.{backend.value}.{profile.value}_wid"
    if importlib.util.find_spec(name):
        return name
    else:
        return None


def get_wid_handler(backend: Backend, profile: Profile):
    """
    Returns a WID handler that searches for hdl_wid_<N> in:
      1. autopts.ptsprojects.<backend>.<profile>_wid  (if the module exists)
      2. autopts.wid.<profile>

    For GATT profiles the client vs. server WID module is chosen at call time
    based on the active stack services and the test-case name.
    """
    def handler(wid, description, test_case_name):
        from autopts.wid import generic_wid_hdl
        log("%r.%r handler, wid=%r, tc=%r", backend.value, profile.value, wid, test_case_name)

        if profile in (Profile.GATT, Profile.GATTC):
            stack = get_stack()
            if stack.is_svc_supported("GATT_CL") and "GATT/CL" in test_case_name:
                gatt_profile = Profile.GATT_CLIENT
            else:
                gatt_profile = Profile.GATT
            backend_mod = _backend_wid_module(backend, gatt_profile)
            ns = ([backend_mod] if backend_mod else []) + [f"autopts.wid.{gatt_profile.value}"]
            return generic_wid_hdl(wid, description, test_case_name, ns)

        backend_mod = _backend_wid_module(backend, profile)
        ns = ([backend_mod] if backend_mod else []) + [f"autopts.wid.{profile.value}"]
        return generic_wid_hdl(wid, description, test_case_name, ns)

    return handler
