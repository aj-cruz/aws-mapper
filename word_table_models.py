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
                        "paragraphs": [{"style": "regularbold", "text": "ENI"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "SUBNETS"}]
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

endpoint_services_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            { # Row 0
                "cells": [
                    { # Cell 0
                        "background": orange_spacer,
                        "paragraphs": [] # This will be the Endpoint Service Name (proper name, not the tag)
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
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    },
                ]
            },
            { # Row 1
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    { # Cell 3
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    },
                ]
            },
            { # Row 2
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TYPE"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AVAILABILITY ZONES"}]
                    },
                    { # Cell 3
                        "paragraphs": []
                    },
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    },
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
            { # Row 0
                "cells": [
                    {
                        "background": orange_spacer,
                        "paragraphs": [{"style": "regularbold", "text": "TRANSIT GATEWAY"}]
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
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    }
                ]
            },
            { # Row 1
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TGW NAME"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "merge": None
                    },
                    { # Cell 3
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ID"}]
                    },
                    { # Cell 4
                        "paragraphs": []
                    },
                    { # Cell 5
                        "merge": None
                    }
                ]
            },
            { # Row 2
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "BGP ASN"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "merge": None
                    },
                    { # Cell 3
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "OWNER ID"}]
                    },
                    { # Cell 4
                        "paragraphs": []
                    },
                    { # Cell 5
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
            "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE ASSOC"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "SUBNETS"}]
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
        {
            "merge": None
        }
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
            { # Row 0
                "cells": [
                    { # Cell 0
                        "background": orange_spacer,
                        "paragraphs": [{"style": "regularbold", "text": ""}]
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
                    {
                        "merge": None
                    },
                    {
                        "merge": None
                    },
                ]
            },
            { # Row 1
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "NAME"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AMI"}]
                    },
                    { # Cell 3
                        "paragraphs": []
                    },
                    { # Cell 4
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INST TYPE / ENA SUPPORT"}]
                    },
                    { # Cell 5
                        "paragraphs": []
                    }
                ]
            },
            { # Row 2
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "AZ"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PRIV IP"}]
                    },
                    { # Cell 3
                        "paragraphs": []
                    },
                    { # Cell 4
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PUB IP"}]
                    },
                    { # Cell 5
                        "paragraphs": []
                    }
                ]
            },
            { # Row 3
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "PLATFORM"}]
                    },
                    { # Cell 1
                        "paragraphs": []
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ARCHITECTURE"}]
                    },
                    { # Cell 3
                        "paragraphs": []
                    },
                    { # Cell 4
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "STATE"}]
                    },
                    { # Cell 5
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
            "paragraphs": [{"style": "regularbold", "text": "INT ID/DESCRIPTION"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "PRIV IPS"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "PUB IP"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "SUBNET"}]
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

lb_target_group_tbl = {
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
                        "paragraphs": [{"style": "regularbold", "text": "PROTOCOL:PORT"}]
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TARGET TYPE"}]
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
                        "paragraphs": [{"style": "regularbold", "text": "LOAD BALANCER ARNS"}]
                    },
                    {
                        "paragraphs": []
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
                    {
                        "merge": None
                    }
                ]
            },
            {
                "cells": [
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "HEALTH CHECK PROTOCOL"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "HEALTH CHECK PORT"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "HEALTHY THRESHOLD"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "UNHEALTHY THRESHOLD"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TIMEOUT SECONDS"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "INTERVAL SECONDS"}]
                    },
                ]
            },
            {
                "cells": [
                    {
                        "paragraphs": []
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "paragraphs": []
                    },
                    {
                        "paragraphs": []
                    },
                ]
            },
            {
                "cells": [
                    {
                        "background": green_spacer,
                        "paragraphs": [{"style": "regularbold", "text": "TARGETS AND TARGET HEALTH"}]
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
                        "paragraphs": [{"style": "regularbold", "text": "TARGET ID"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TARGET PORT"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "STATE"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REASON"}]
                    },
                    {
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "DESCRIPTION"}]
                    },
                    {
                        "merge": None
                    }
                ]
            }
        ]
    }
}

best_practices_tbl = {
		"table": {
			"style": "plain",
			"rows": [
                { # Row 0
                    "cells": [
                        { # Cell 0
                            "background": None,
                            "paragraphs": [{"style": "regularbold", "text": ""}]
                        },
                        {"merge": None},
                        {"merge": None},
                        {"merge": None},
                        {"merge": None}
                    ]
                },
                { # Row 1
                    "cells": [
                        { # Cell 0
                            "background": table_header_color,
                            "paragraphs": [{"style": "regularbold", "text": "TEST DESCRIPTION"}]
                        },
                        { # Cell 1
                            "paragraphs": [{"style": "No Spacing", "text": ""}]
                        },
                        {"merge": None},
                        {"merge": None},
                        {"merge": None}
                    ]
                },
                { # Row 2
                    "cells": [
                        { # Cell 0
                            "background": table_header_color,
                            "paragraphs": [{"style": "regularbold", "text": "RESULT DETAILS"}]
                        },
                        { # Cell 1
                            "paragraphs": [{"style": "No Spacing", "text": ""}]
                        },
                        {"merge": None},
                        {"merge": None},
                        {"merge": None}
                    ]
                }
            ]
		}
	}

account_dashboard_tbl = {
    "table": {
        "style": "plain",
        "rows": [
            { # Row 0
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACCOUNT ID"}]
                    },
                    { # Cell 1
                        "paragraphs": [{"style": "No Spacing", "text": ""}]
                    },
                    { # Cell 2
                        "merge": None
                    },
                    { # Cell 3
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "ACCOUNT ALIAS"}]
                    },
                    { # Cell 4
                        "paragraphs": [{"style": "No Spacing", "text": ""}]
                    },
                    { # Cell 5
                        "merge": None
                    }
                ]
            },
            { # Row 1
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "REGIONS IN USE"}]
                    },
                    { # Cell 1
                        "paragraphs": [{"style": "No Spacing", "text": ""}]
                    },
                    { # Cell 2
                        "merge": None
                    },
                    { # Cell 3
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "EC2/VPC COUNT"}]
                    },
                    { # Cell 4
                        "paragraphs": [{"style": "No Spacing", "text": ""}]
                    },
                    { # Cell 5
                        "merge": None
                    }
                ]
            },
            { # Row 2
                "cells": [
                    { # Cell 0
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "TRANSIT GATEWAY ANALYSIS"}]
                    },
                    { # Cell 1
                        "merge": None
                    },
                    { # Cell 2
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPN ANALYSIS"}]
                    },
                    { # Cell 3
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "VPC ANALYSIS"}]
                    },
                    { # Cell 4
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "LOAD BALANCER ANALYSIS"}]
                    },
                    { # Cell 5
                        "background": table_header_color,
                        "paragraphs": [{"style": "regularbold", "text": "EC2 INST ANALYSIS"}]
                    },
                ]
            }
        ]
    }
}