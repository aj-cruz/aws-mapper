from map import table_header_color

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
            "paragraphs": []
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "ROUTE TABLE ID"}]
        },
        {
            "paragraphs": []
        }
    ]
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

vgw_cgw_tbl_header = {
    "cells": [
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "CGW NAME"}]
        },
        {
            "background": table_header_color,
            "paragraphs": [{"style": "regularbold", "text": "CGW ID"}]
        },
        {
            "merge": None
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
    ]
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
