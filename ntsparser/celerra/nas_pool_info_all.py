"""nas pool info all Sheet"""
import textwrap
from collections import namedtuple
from typing import Any

from ntsparser.formatting import (
    build_header, column_format, set_cell_to_number, style_value_cell)
from ntsparser.parsing import get_parser_header, run_parser_over
from ntsparser.utils import sheet_process_output


NAS_POOL_INFO_TMPL = textwrap.dedent("""\
    Value Required,Filldown Hostname (\S+)
    Value Required ID (.+)
    Value Name (.+)
    Value Description (.+)
    Value ACL (.+)
    Value InUse (.+)
    Value Clients (.+)
    Value Members (.+)
    Value StorageSystems (.+)
    Value DefaultSliceFlag (.+)
    Value IsUserDefined (.+)
    Value Thin (.+)
    Value VirtuallyProvisioned (.+)
    Value DiskType (.+)
    Value ServerVisibility (.+)
    Value VolumeProfile (.+)
    Value IsDynamic (.+)
    Value IsGreedy (.+)
    Value NumStripeMembers (.+)
    Value StripeSize (.+)

    Start
      ^\s*Output from:\s+/bin/hostname -> HostLine

    HostLine
      ^\s*${Hostname}\s*$$ -> NASPoolStart

    NASPoolStart
      ^\s*Output from:\s+/nas/bin/nas_pool\s+-info\s+-all -> NASPoolLines

    NASPoolLines
      ^\s*id\s+=\s+${ID}
      ^\s*name\s+=\s+${Name}
      ^\s*description\s+=\s+${Description}
      ^\s*acl\s+=\s+${ACL}
      ^\s*in_use\s+=\s+${InUse}
      ^\s*clients\s+=\s+${Clients}
      ^\s*members\s+=\s+${Members}
      ^\s*storage_system\(s\)\s+=\s+${StorageSystems}
      ^\s*default_slice_flag\s+=\s+${DefaultSliceFlag}
      ^\s*is_user_defined\s+=\s+${IsUserDefined}
      ^\s*thin\s+=\s+${Thin}
      ^\s*virtually_provisioned\s*=${VirtuallyProvisioned}
      ^\s*disk_type\s+=\s+${DiskType}
      ^\s*server_visibility\s+=\s+${ServerVisibility}
      ^\s*volume_profile\s+=\s+${VolumeProfile}
      ^\s*is_dynamic\s+=\s+${IsDynamic}
      ^\s*is_greedy\s+=\s+${IsGreedy}
      ^\s*num_stripe_members\s+=\s+${NumStripeMembers}
      ^\s*stripe_size\s+=\s+${StripeSize} -> Record NASPoolLines
      ^\s*(\*+) -> Start
""")


def process(workbook: Any, content: str) -> None:
    """Process nas pool info all worksheet

    :param workbook:
    :param content:
    """
    worksheet = workbook.get_sheet_by_name('nas pool info all')

    headers = get_parser_header(NAS_POOL_INFO_TMPL)

    RowTuple = namedtuple('RowTuple', headers)  # pylint: disable=invalid-name

    build_header(worksheet, headers)

    nas_pool_out = run_parser_over(content, NAS_POOL_INFO_TMPL)

    final_col, final_row = 0, 0
    for row_n, row_tuple in enumerate(map(RowTuple._make, nas_pool_out), 2):
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
        'NASPoolInfoAllTable',
        'nas pool info all',
        final_col,
        final_row)
