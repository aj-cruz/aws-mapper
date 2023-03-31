import argparse, os, sys, json, pathlib
from rich import print as rprint
from rich import print_json as jprint

# GLOBAL VARIABLES
log_file = "results.json"
topology_file_path = "./topologies"

# SET COMMANDLINE ARGUMENT PARAMETERS
parser = argparse.ArgumentParser()
# QUERY ARGUMENTS
parser.add_argument(
    '--getinstancesinsubnet',
    action='store',
    dest='subnet',
    default=None,
    metavar="subnet-<subnet_id>",
    help='AWS Subnet ID (subnet-xxxxxxxxxx). This will return a list of all instances in the provided subnet'
    )
parser.add_argument(
    '--getinstancesinvpc',
    action='store',
    dest='instvpc',
    default=None,
    metavar="vpc-<vpc_id>",
    help='AWS VPC ID (vpc-xxxxxxxxxx). This will return a list of all instances in the provided VPC'
    )
parser.add_argument(
    '--getinstancebyeip',
    action='store',
    dest='instbyeip',
    default=None,
    metavar="<elastic_ip_address>",
    help='AWS Elastic IP Address. This will return the instance attached to the EIP'
    )
parser.add_argument(
    '--getsubnetbyname',
    action='store',
    dest='subname',
    default=None,
    metavar="<subnet_name>",
    help='AWS Subnet Name. This will return the subnet ID'
    )
parser.add_argument(
    '--locatevpc',
    action='store',
    dest='locvpc',
    default=None,
    metavar="vpc-<vpc_id>",
    help='AWS VPC ID (vpc-xxxxxxxxxx). This will return the account id and json file where the VPC is found.'
    )
parser.add_argument(
    '-t',
    '--topology',
    action='store',
    dest='topology_file',
    default=None,
    metavar="<filename>",
    help='The name of the topology file (JSON) to read from.'
    )
# FLAGS
parser.add_argument(
    '-l',
    '--log',
    action='store_true',
    default=False,
    dest='enable_log',
    help=f'Enable logging to: {os.getcwd()}/{log_file}'
    )
args = parser.parse_args()
# Pull out and store just the queries from all possible arguments
query = [{k:v} for k, v in vars(args).items() if v and not isinstance(v, bool) and not k == "topology_file"]

# HELPER FUNCTIONS
def get_name_from_tags(tags):
    try:
        name = [tag['Value'] for tag in tags if tag['Key'] == "Name"][0]
    except KeyError:
        name = ""
    except IndexError:
        name = ""
    return name

def build_topology_file_list():
    # Build a list of json files in the topologies directory
    fp = pathlib.Path(topo_fp)
    return [f.name for f in fp.iterdir() if f.is_file() and f.name.endswith(".json")]

# FUNCTIONS
def validate_args():
    if len(query) > 1:
        rprint("\n\n:x: [red]Multiple queries are not presently supported. Please re-run with a single query.\n\n")
        sys.exit(1)

    if not list(query[0].keys())[0] in ["locvpc", "instbyeip"] and not args.topology_file.endswith(".json"):
        rprint(f"\n\n:x: [red]File [blue]{args.topology_file} [red]does not appear to be a JSON file (extension not .json). The script requires a JSON file.\n\n")
        sys.exit(1)

    if list(query[0].keys())[0] == "subnet":
        if not query[0]['subnet'].startswith("subnet-"):
            rprint(f"\n\n:x: [red]The provided subnet [purple]{query[0]['subnet']} [red]does not appear to be a valid AWS Subnet ID.")
            rprint("[red]AWS Subnet IDs should begin with 'subnet-'\n\n")
            sys.exit(1)
    elif list(query[0].keys())[0] == "instvpc":
        if not query[0]['instvpc'].startswith("vpc-"):
            rprint(f"\n\n:x: [red]The provided vpc [purple]{query[0]['instvpc']} [red]does not appear to be a valid AWS VPC ID.")
            rprint("[red]AWS VPC IDs should begin with 'vpc-'\n\n")
            sys.exit(1)
    elif list(query[0].keys())[0] == "locvpc":
        if not query[0]['locvpc'].startswith("vpc-"):
            rprint(f"\n\n:x: [red]The provided vpc [purple]{query[0]['locvpc']} [red]does not appear to be a valid AWS VPC ID.")
            rprint("[red]AWS VPC IDs should begin with 'vpc-'\n\n")
            sys.exit(1)

def build_topology_file_path():
    if topology_file_path.startswith("./"): # using relative path so strip the "." and prepend the current working directory
        topology_fp = os.getcwd() + topology_file_path.replace("./", "/")
    else:
        topology_fp = topology_file_path
    if not topology_fp.endswith("/"):
        topology_fp += "/"
    return topology_fp

def read_topology_from_json():
    try:
        if args.topology_file.startswith("./"): # using relative path so strip the "." and prepend the current working directory
            topology_file = topo_fp + args.topology_file.replace("./", "")
        else:
            topology_file = topo_fp + args.topology_file
        with open(f"{topology_file}", "r") as f:
            topology = json.load(f)
    except FileNotFoundError:
        rprint(f"\n\n:x: [red]No such file or directory: [blue]{os.getcwd()}{args.topology_file}\n\n")    
        sys.exit(1)

    return topology

def run_instances_by_subnet_query():
    results = [{
        "region": k,
        "vpc_id": inst['VpcId'],
        "name": get_name_from_tags(inst['Tags']),
        "inst_id": inst['InstanceId']
        } for k, v in topology.items() if isinstance(v, dict) and "vpcs" in v.keys() and v['vpcs'] for vpc in v['vpcs'] if vpc['ec2_instances'] for inst in vpc['ec2_instances'] if inst['SubnetId'] == query['value'].lower()]
    jprint(data=results)
    rprint(f"\n {len(results)} [yellow]Instances found in subnet: {query['value']}")
    if args.enable_log:
        with open(log_file, "w") as f:
            f.write(json.dumps(results,indent=4))

def run_instances_by_vpc_query():
    results = [{
        "region": k,
        "vpc_id": inst['VpcId'],
        "name": get_name_from_tags(inst['Tags']),
        "inst_id": inst['InstanceId']
        } for k, v in topology.items() if isinstance(v, dict) and "vpcs" in v.keys() and v['vpcs'] for vpc in v['vpcs'] if vpc['ec2_instances'] and vpc['VpcId'] == query['value'].lower() for inst in vpc['ec2_instances']]
    jprint(data=results)
    rprint(f"\n {len(results)} [yellow]Instances found in VPC: {query['value']}")
    if args.enable_log:
        with open(log_file, "w") as f:
            f.write(json.dumps(results,indent=4))

def run_subnet_by_name_query():
    try:
        results = [sub['SubnetId'] for k, v in topology.items() if isinstance(v, dict) and "vpcs" in v.keys() and v['vpcs'] for vpc in v['vpcs'] if vpc['subnets'] for sub in vpc['subnets'] if get_name_from_tags(sub['Tags']) == query['value']][0]
        rprint(f"\n\n{query['value']} = {results}")
    except IndexError:
        rprint(f"\n\n[red]Subnet '{query['value']}' not found in file: [blue]{args.topology_file}\n")

def run_locate_vpc():
    file_list = build_topology_file_list()
    
    for f in file_list:
        with open(topo_fp + f, "r") as f:
            topology = json.load(f)
        for k, v in topology.items():
            if isinstance(v, dict) and "vpcs" in v.keys():
                for vpc in v['vpcs']:
                    if vpc['VpcId'] == query['value']:
                        try: # Get VPC Name
                            vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                        except KeyError:
                            vpc_name = ""
                        except IndexError:
                            vpc_name = ""
                        rprint(f"\n[yellow]{query['value']} found in Account: [white] {topology['account']['id']}")
                        rprint(f"[yellow]VPC Name: [white]{vpc_name}")
                        rprint(f"[yellow]File: [blue]{f.name}")

def run_instance_by_eip():
    file_list = build_topology_file_list()
    
    for f in file_list:
        with open(topo_fp + f, "r") as f:
            topology = json.load(f)
        for k, v in topology.items():
            if isinstance(v, dict) and "vpcs" in v.keys():
                inst = [{
                    "region": k,
                    "vpc_id": vpc['VpcId'],
                    "name": get_name_from_tags(inst['Tags']),
                    "inst_id": inst['InstanceId']
                } for vpc in v['vpcs'] for inst in vpc['ec2_instances'] for intf in inst['NetworkInterfaces'] if "Association" in intf.keys() and intf['Association']['PublicIp'] == query['value']]
                if len(inst) > 0:
                    jprint(data=inst)

if __name__ == "__main__":
    validate_args()
    topo_fp = build_topology_file_path()
    query = [{"type":k,"value":v} for k, v in query[0].items()][0]
    if not query['type'] in ["locvpc", "instbyeip"]:
        topology = read_topology_from_json()
    
    if query['type'] == "subnet":
        run_instances_by_subnet_query()
    elif query['type'] == "subname":
        run_subnet_by_name_query()
    elif query['type'] == "instvpc":
        run_instances_by_vpc_query()
    elif query['type'] == "locvpc":
        run_locate_vpc()
    elif query['type'] == "instbyeip":
        run_instance_by_eip()