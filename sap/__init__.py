"""
SAP 自动化操作模块

使用方法:
    from sap import Sap, SapConfig, OrderData, RevenueData, OperationFlags, HourData, SapResult
"""

from sap.models import (
    SapConfig,
    OrderData,
    RevenueData,
    OperationFlags,
    HourData,
    SapResult,
)
from sap.function import Sap

__all__ = [
    'Sap',
    'SapConfig',
    'OrderData',
    'RevenueData',
    'OperationFlags',
    'HourData',
    'SapResult',
]
