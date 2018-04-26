"""Gets clusters from each input file(XtremIO)"""
import textwrap
from supergrep.parsing import run_parser_over

SHOW_CLUSTERS_INFO_TMPL = textwrap.dedent("""\
    Value Required ClusterName (\S+)

    Start
      ^\s*Cluster-Name -> ClusterInfoLine

    ClusterInfoLine
      ^\s*${ClusterName}\s+(.+) -> Record
""")

def process(content: str) -> list:
    """Process clusters (XtremIO)

    :param content:
    :return:
    """
    clusters_info_out = run_parser_over(content, SHOW_CLUSTERS_INFO_TMPL)
    clusters_data = [
        ("Cluster Names:",(
            ', '.join(cluster[0] for cluster in clusters_info_out)), '', '')
    ]

    return clusters_data