from map import table_header_color, orange_spacer, green_spacer, red_spacer

parent_tbl = {
		"table": {
			"style": None,
			"rows": []
		}
	}

vpc_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REGION"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC CIDR"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INSTANCE QTY"}]
                    }
                ]
            }
        ]
    }
}

rt_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RT NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RT ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ROUTE QTY "}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUB ASSOC QTY"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "EDGE ASSOC QTY"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ROUTE PROPAGATIONS"}]
                    },
                ]
            }
        ]
    }
}

rt_routes_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DESTINATION"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TARGET"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ORIGIN"}]
                    }
                ]
            }
        ]
    }
}

subnet_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNET"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNET NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AVAILABILITY ZONE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NET ACL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INST QTY"}]
                    }
                ]
            }
        ]
    }
}

prefix_list_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "CIDRS"}]
                    }
                ]
            }
        ]
    }
}

netacls_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACL NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACL ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "IS DEFAULT"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNETS"}]
                    }
                ]
            }
        ]
    }
}

netacl_in_entries_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RULE NUMBER"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PROTOCOL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PORT RANGE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SOURCE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACTION"}]
                    }
                ]
            }
        ]
    }
}

netacl_out_entries_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RULE NUMBER"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PROTOCOL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PORT RANGE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SOURCE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACTION"}]
                    }
                ]
            }
        ]
    }
}

sec_grps_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SG NAME (TAG)"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SG NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SG ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DESCRIPTION"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "IN/OUT RULE COUNT"}]
                    }
                ]
            }
        ]
    }
}

sec_grp_in_entries_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PROTOCOL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PORT RANGE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SOURCE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DESCRIPTION"}]
                    }
                ]
            }
        ]
    }
}

sec_grp_out_entries_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PROTOCOL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PORT RANGE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SOURCE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DESCRIPTION"}]
                    }
                ]
            }
        ]
    }
}

igw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "IGW NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    }
                ]
            }
        ]
    }
}

eigw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "EIGW NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    }
                ]
            }
        ]
    }
}

ngw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NGW NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNET"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    }
                ]
            }
        ]
    }
}

ngw_address_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PUBLIC IP"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PRIVATE IP"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NET INTERFACE ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "IS PRIMARY"}]
                    }
                ]
            }
        ]
    }
}

endpoint_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SVC NAME"}]
                    }
                ]
            }
        ]
    }
}

vpc_peering_requester_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "CONNECTION ID"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REQUESTER REGION"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REQUESTER VPC"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REQUESTER CIDR BLOCK"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REQUESTER OWNER ID"}]
                    }
                ]
            }
        ]
    }
}

vpc_peering_accepter_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ACCEPTER REGION"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ACCEPTER VPC"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ACCEPTER CIDR BLOCK"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ACCEPTER OWNER ID"}]
        }
    ]
}

tgw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TGW NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "BGP ASN"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "OWNER ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    }
                ]
            }
        ]
    }
}

tgw_attachment_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ATTACHMENT NAME"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ATTACHMENT ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "RESOURCE ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE ASSOC."}]
        }
    ]
}

tgw_rt_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE NAME"}]
        },
        {
            "merge": None
        },
        {
            "merge": None
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE ID"}]
        },
        {
            "merge": None
        },
    ]
}

tgw_routes_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "CIDR"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RESOURCE TYPE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "RESOURCE ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ATTACHMENT ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ROUTE TYPE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ROUTE STATE"}]
                    }
                ]
            }
        ]
    }
}

vpn_cgw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DEV NAME"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "IP ADDRESS"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "BGP ASN"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                ]
            }
        ]
    }
}

vpn_tgw_conn_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "CONN NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPN ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TGW ID"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "CGW ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "LOCAL CIDR"}]
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REMOTE CIDR"}]
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "OUTSIDE ADDRESS TYPE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INSIDE IP VERSION"}]
                    },
                ]
            }
        ]
    }
}

vgw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "GATEWAY NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "GW ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "BGP ASN"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC ATTACHMENT"}]
                    },
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    }
                ]
            }
        ]
    }
}

vgw_conn_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "CONNECTION NAME"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "CGW ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "LOCAL CIDR"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "REMOTE CIDR"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "OUTSIDE ADDRESS TYPE"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "INSIDE IP VERSION"}]
        },
    ]
}

vgw_conn_tunnel_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "OUTSIDE IP"}]
        },
        {
            "merge": None
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "INSIDE CIDR"}]
        },
        {
            "merge": None
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "IPSEC STATUS"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "TUN STATUS"}]
        },
    ]
}

ec2_inst_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AMI"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INST TYPE"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AZ"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PRIV IP"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PUB IP"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PLATFORM"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ARCHITECTURE"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "STATE"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            }
        ]
    }
}

ec2_inst_interface_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "INT ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "PRIV IP"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "PUB IP"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "SUB ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "SG IDS"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "DEV INDEX"}]
        },
    ]
}

dcgw_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": orange_spacer,
                        "paragraphs": [{"style": "regularbold", "text": "DIRECT CONNECT GATEWAY"}]
                    },
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "GW NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "GW ID"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AWS ASN"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            }
        ]
    }
}

dcgw_conn_rows = [
    {
        "cells": [
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "NAME"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "ID"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "REGION"}]
            },
            {
                "paragraphs": []
            }
        ]
    },
    {
        "cells": [
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "LOCATION"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "PARTNER"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "BANDWIDTH"}]
            },
            {
                "paragraphs": []
            }
        ]
    },
    {
        "cells": [
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "JUMBO FRAMES"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "MACSEC CAPABLE"}]
            },
            {
                "paragraphs": []
            },
            {
                "background": table_header_color,
                "paragraphs": [{"style": "regularbold", "text": "ENCRYPTION STATUS"}]
            },
            {
                "paragraphs": []
            }
        ]
    }
]

dcgw_vif_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "NAME"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ID"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "AMAZON ADDRESS"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "MTU"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "BGP PEER:STATUS"}]
        },
    ]
}

load_balancer_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            {
                "cells": [
                    {
                        "background": orange_spacer,
                        "paragraphs": []
                    },
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "STATE"}]
                    },
                    {
                        "paragraphs": []
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": green_spacer,
                        "paragraphs": [{"style": "regularbold", "text": "NETWORK MAPPINGS"}]
                    },
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                    {"merge": None},
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AVAILABILITY ZONE"}]
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNET"}]
                    },
                    {
                        "merge": None
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ADDRESSES"}]
                    },
                    {
                        "merge": None
                    }
                ]
            }
        ]
    }
}

load_balancer_listener_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "PROTOCOL:PORT"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "FORWARD GROUP"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "LISTENER ARN"}]
        },
        {
            "merge": None
        },
        {
            "merge": None
        },
        {
            "merge": None
        },
    ]
}