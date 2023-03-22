import os, json, pathlib, platform
from rich import print as rprint
from rich import print_json as jprint

# GLOBAL VARIABLES
topology_file_path = "./topologies"

# FUNCTIONS
def build_topology_file_path():
    if topology_file_path.startswith("./"): # using relative path so strip the "." and prepend the current working directory
        topology_fp = os.getcwd() + topology_file_path.replace("./", "/")
    else:
        topology_fp = topology_file_path
    if not topology_fp.endswith("/"):
        topology_fp += "/"
    return topology_fp

def get_list_of_topology_files():
    fp = pathlib.Path(topo_fp)
    return [f.name for f in fp.iterdir() if f.is_file() and f.name.endswith(".json")]

def slasher():
    # Returns the correct file system slash for the detected platform
    return "\\" if system_os == "windows" else "/"

if __name__ == "__main__":
    system_os = platform.system().lower()
    topo_fp = build_topology_file_path()
    topo_files = get_list_of_topology_files()

    statistics = {
        "accounts": 0,
        "vpcs": 0,
        "route_tables": 0,
        "vpc_routes": 0,
        "prefix_lists": 0,
        "subnets": 0,
        "network_acls": 0,
        "net_acl_in_entries": 0,
        "net_acl_out_entries": 0,
        "security_groups": 0,
        "sg_in_entries": 0,
        "sg_out_entries": 0,
        "internet_gateways": 0,
        "egress_only_internet_gateways": 0,
        "nat_gateways": 0,
        "peering_connections": 0,
        "transit_gateways": 0,
        "transit_gateway_attachments": 0,
        "customer_gateways": 0,
        "vpn_connections": 0,
        "virtual_private_gateways": 0,
        "ec2_instances": 0,
        "ec2_groups": 0,
        "direct_connects": 0
    }
    for file in topo_files:
        with open(f"{topo_fp}{slasher()}{file}", "r") as f:
            topology = json.load(f)
        statistics['accounts'] += 1
        statistics['peering_connections'] += len(topology['vpc_peering_connections'])
        statistics['direct_connects'] += len(topology['direct_connect'])
        for region, attributes in topology.items():
            if isinstance(attributes, dict) and "vpcs" in attributes.keys():
                statistics['vpcs'] += len(attributes['vpcs'])
                for vpc in attributes['vpcs']:
                    statistics['route_tables'] += len(vpc['route_tables'])
                    statistics['vpc_routes'] += len([route for rt in vpc['route_tables'] for route in rt['Routes']])
                    statistics['subnets'] += len(vpc['subnets'])
                    statistics['network_acls'] += len(vpc['network_acls'])
                    statistics['net_acl_in_entries'] += len([entry for acl in vpc['network_acls'] for entry in acl['Entries'] if not entry['Egress']])
                    statistics['net_acl_out_entries'] += len([entry for acl in vpc['network_acls'] for entry in acl['Entries'] if entry['Egress']])
                    statistics['security_groups'] += len(vpc['security_groups'])
                    statistics['sg_in_entries'] += len([entry for sg in vpc['security_groups'] for entry in sg['IpPermissions']])
                    statistics['sg_out_entries'] += len([entry for sg in vpc['security_groups'] for entry in sg['IpPermissionsEgress']])
                    statistics['internet_gateways'] += len(vpc['internet_gateways'])
                    statistics['egress_only_internet_gateways'] += len(vpc['egress_only_internet_gateways'])
                    statistics['nat_gateways'] += len(vpc['nat_gateways'])
                    statistics['virtual_private_gateways'] += len(vpc['vpn_gateways'])
                    statistics['ec2_instances'] += len(vpc['ec2_instances'])
                    statistics['ec2_groups'] += len(vpc['ec2_groups'])
                    statistics['ec2_groups'] += len(vpc['ec2_groups'])
                if "prefix_lists" in attributes.keys():
                    statistics['prefix_lists'] += len(attributes['prefix_lists'])
                statistics['transit_gateways'] += len(attributes['transit_gateways'])
                statistics['transit_gateway_attachments'] += len([atch for tgw in attributes['transit_gateways'] for atch in tgw['attachments']])
                if "customer_gateways" in attributes.keys():
                    statistics['customer_gateways'] += len(attributes['customer_gateways'])
                if "vpn_tgw_connections" in attributes.keys():
                    statistics['vpn_connections'] += len(attributes['vpn_tgw_connections'])

    jprint(data=statistics)
    total_objects = 0
    for each in statistics.values():
        total_objects += each
    rprint(f"[yellow]TOTAL OBJECTS: {total_objects}\n")