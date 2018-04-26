"""Pool Configuration Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from supergrep.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from supergrep.parsing import get_parser_header, run_parser_over
from supergrep.utils import sheet_process_output


POOL_CONFIG_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required VPSupported (.+)
    Value MaximumPools (.+)
    Value MaximumDisksPerPools (.+)
    Value MaximumDisksAllPools (.+)
    Value MaximumPoolLuns (.+)
    Value NumberOfPools (.+)
    Value NumberOfPoolLuns (.+)
    Value MinimumPoolLunSize (.+)
    Value MaximumPoolLunSize (.+)
    Value NumberOfThinLuns (.+)
    Value NumberOfNonThinLuns (.+)
    Value NumberOfDisksUsed (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> LicenseStart

    LicenseStart
      ^\s*Output from:\s+/nas/bin/nas_storage\s+-info\s+-all\s+-o\s+sync=no -> PoolStart

    PoolStart
      ^\s*Pool Configuration -> PoolLines
      
    PoolLines
      ^\s*is VP supported\s+=\s+${VPSupported}
      ^\s*maximum pools\s+=\s+${MaximumPools}
      ^\s*maximum disks per pools\s+=\s+${MaximumDisksPerPools}
      ^\s*maximum disks all pools\s+=\s+${MaximumDisksAllPools}
      ^\s*maximum pool luns\s+=\s+${MaximumPoolLuns}
      ^\s*number of pools\s+=\s+${NumberOfPools}
      ^\s*number of pool luns\s+=\s+${NumberOfPoolLuns}
      ^\s*minimum pool lun size\s+=\s+${MinimumPoolLunSize}
      ^\s*maximum pool lun size\s+=\s+${MaximumPoolLunSize}
      ^\s*number of thin luns\s+=\s+${NumberOfThinLuns}
      ^\s*number of non-thin luns\s+=\s+${NumberOfNonThinLuns}
      ^\s*number of disks used\s+=\s+${NumberOfDisksUsed} -> Record
      ^\s*Storage Groups\s*$$ -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process Pool Configuration worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('Pool Configuration')

    headers = get_parser_header(POOL_CONFIG_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    pool_config_out = run_parser_over(content, POOL_CONFIG_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, pool_config_out), 2):
        for col_n, col_value in \
                enumerate(row_tuple._asdict().values(), ord('A')):
            cell = worksheet['{}{}'.format(column_format(col_n), row_n)]
            cell.value = str.strip(col_value)
            style_value_cell(cell)
            set_cell_to_number(cell)
            final_col = col_n
        final_row = row_n

    sheet_process_output(
        worksheet,
        'PoolConfigurationTable',
        'Pool Configuration',
        final_col,
        final_row)
