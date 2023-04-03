import boto3, botocore.exceptions, requests, sys, datetime, json, os, argparse, pathlib, datetime, platform
from rich import print as rprint
from rich import print_json as jprint
from docx import Document
from dcnet_msofficetools.docx_extensions import build_table, replace_placeholder_with_table
from copy import deepcopy
import word_table_models, word_title_page
from urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

# SET COMMANDLINE ARGUMENT PARAMETERS
parser = argparse.ArgumentParser()
parser.add_argument(
    '-t',
    action='store_true',
    default=False,
    dest='skip_topology',
    help='Skip bulding topology file from AWS API (it already exists in working directory)'
    )
args = parser.parse_args()

# GLOBAL VARIABLES
output_verbosity = 0   # 0 (Default) or 1 (Verbose)
topology_folder = "topologies"
word_template = "template.docx"
table_header_color = "506279"
green_spacer = "8FD400" # CC Green/Lime
red_spacer = "F12938" # CC Red
orange_spacer = "FF7900" # CC Orange
alternating_row_color = "D5DCE4"
region_list = [] # Leave blank to auto-pull and check all regions
aws_protocol_map = { # Maps AWS protocol numbers to user-friendly names
    "-1": "All Traffic",
    "6": "TCP",
    "17": "UDP",
    "1": "ICMPv4",
    "58": "ICMPv6"
}
non_region_topology_keys = ["account", "vpc_peering_connections", "direct_connect_gateways"]

# HELPER FUNCTIONS
def datetime_converter(obj):
    # Converts datetime objects to a string timestamp. Needed to render JSON
    if isinstance(obj, datetime.datetime):
        return obj.__str__()

def create_word_obj_from_template(tfile):
    try:
        return Document(tfile)
    except:
        rprint(f"\n\n:x: [red]Could not open [blue]{tfile}[red]. Please make sure it exists and is a valid Microsoft Word document. Exiting...")
        sys.exit(1)

def extract_name_from_aws_tags(tags):
    # Extract the name from a list of AWS tags
    try:
        name = [tag['Value'] for tag in tags if tag['Key'] == "Name"][0]
    except KeyError:
        # Object has no name
        name = "<unnamed>"
    except IndexError:
        # Object has no name
        name = "<unnamed>"
    return name

def get_subnet_name_by_id(source_subnet_id, vpc):
    # Return a subnet's name tag (if available) by looking up the name in the topology VPC structure
    for subnet in vpc['subnets']:
        if subnet['SubnetId'] == source_subnet_id:
            return extract_name_from_aws_tags(subnet['Tags'])
# MAIN FUNCTIONS
def get_regions():
    if region_list:
        return region_list
    else:
        response = ec2.describe_regions()
        discovered_regions = [x['RegionName'] for x in response['Regions']]
        return discovered_regions

def add_regions_to_topology():
    for region in available_regions:
        topology[region] = {
            "vpcs": [],
            "transit_gateways": [],
            "instances": []
        }

def fingerprint_vpc(region, vpc, ec2):
    '''
    Fingerprint/interrogate a VPC to see if it is a default & unused VPC. These will be filtered from the results.
    FINGERPRINT OF A DEFAULT UNUSED VPC:
        1- IsDefault = true
        2- CIDR Block = 172.31.0.0/16
        3- Three (3) Subnets = 172.31.0.0/20, 172.31.16.0/20, & 172.31.32.0/20
        4- One (1) Route Table with no tags and 2 routes with DestinationCidrBlock = 172.31.0.0/16 & 0.0.0.0/0
        5- One (1) Internet Gateway with no tags
        6- Zero (0) NAT Gateways
        7- One (1) Security Group with description = "default VPC security group"
        8- Zero (0) EC2 Instances
    '''

    fingerprint_checks = []
    if output_verbosity == 1:
        rprint(f"\n[white]Running VPC fingerprint on: {region}\{vpc['VpcId']}...")

    # TEST 1
    status = "pass" if vpc['IsDefault'] else "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]IsDefault = true Check: [green]PASS" if status == "pass" else "    [white]IsDefault = true Check: [red]FAIL"
        rprint(msg)

    # TEST 2
    status = "pass" if vpc['CidrBlock'] == "172.31.0.0/16" else "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]CIDR = 172.31.0.0/16 Check: [green]PASS" if status == "pass" else "    [white]CIDR = 172.31.0.0/16 Check: [red]FAIL"
        rprint(msg)

    # TEST 3
    subnets = [subnet['CidrBlock'] for subnet in ec2.describe_subnets()['Subnets'] if subnet['VpcId'] == vpc['VpcId']]
    if len(subnets) == 3 and "172.31.0.0/20" in subnets and "172.31.16.0/20" in subnets and "172.31.32.0/20" in subnets:
        status = "pass"
    elif len(subnets) == 4 and "172.31.0.0/20" in subnets and "172.31.16.0/20" in subnets and "172.31.32.0/20" in subnets and "172.31.48.0/20" in subnets:
        status = "pass"
    else:
        status = "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]3 Default Subnets Check: [green]PASS" if status == "pass" else "    [white]3 Default Subnets Check: [red]FAIL"
        rprint(msg)

    # TEST 4
    route_tables = [rt for rt in ec2.describe_route_tables()['RouteTables'] if rt['VpcId'] == vpc['VpcId']]
    if len(route_tables) == 1 and not route_tables[0]['Tags']: # Only 1 route table and it has no tags
        routes = [route['DestinationCidrBlock'] for route in route_tables[0]['Routes']]
        if len(routes) == 2 and "0.0.0.0/0" in routes and "172.31.0.0/16" in routes: # Exactly 2 specific routes
            status = "pass"
        else:
            status = "fail"
    else:
        status = "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]Route Table Check: [green]PASS" if status == "pass" else "    [white]Route Table Check: [red]FAIL"
        rprint(msg)

    # TEST 5
    igws = [igw for igw in ec2.describe_internet_gateways()['InternetGateways'] if igw['Attachments'][0]['VpcId'] == vpc['VpcId']] 
    if len(igws) == 1 and not igws[0]['Tags']: # Exactly 1 IGW with no Tags
        status = "pass"
    else:
        status = "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]Internet Gateway Check: [green]PASS" if status == "pass" else "    [white]Internet Gateway Check: [red]FAIL"
        rprint(msg)

    # TEST 6
    natgws = [natgw for natgw in ec2.describe_nat_gateways()['NatGateways'] if natgw['VpcId'] == vpc['VpcId']]
    status = "pass" if len(natgws) == 0 else "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]NAT Gateway Check: [green]PASS" if status == "pass" else "    [white]NAT Gateway Check: [red]FAIL"
        rprint(msg)

    # TEST 7
    sec_grps = [sg for sg in ec2.describe_security_groups()['SecurityGroups'] if sg['VpcId'] == vpc['VpcId']]
    if len(sec_grps) == 1 and sec_grps[0]['Description'].lower() == "default vpc security group":
        status = "pass"
    else:
        status = "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]Security Group Check: [green]PASS" if status == "pass" else "    [white]Security Group Check: [red]FAIL"
        rprint(msg)

    # TEST 8
    reservations = ec2.describe_instances()['Reservations']
    instances = sum([inst['Instances'] for inst in reservations], [])
    this_vpc_instances = [inst for inst in instances if "VpcId" in inst.keys() and inst['VpcId'] == vpc['VpcId']]
    status = "pass" if len(this_vpc_instances) == 0 else "fail"
    fingerprint_checks.append(status)
    if output_verbosity == 1:
        msg = "    [white]EC2 Instance Check: [green]PASS" if status == "pass" else "    [white]EC2 Instance Check: [red]FAIL"
        rprint(msg)

    return False if "fail" in fingerprint_checks else True

def add_vpcs_to_topology():
    for region in topology:
        if not region in non_region_topology_keys:
            rprint(f"    [yellow]Interrogating Region {region} for VPCs...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            elb = boto3.client('elbv2',region_name=region,verify=False)
            try:
                response = ec2.describe_vpcs()['Vpcs']
                topology[region]["non_vpc_lb_target_groups"] = [tg for tg in elb.describe_target_groups()['TargetGroups'] if not "VpcId" in tg.keys()]
                for tg in topology[region]["non_vpc_lb_target_groups"]:
                    tg['HealthChecks'] = elb.describe_target_health(TargetGroupArn=tg['TargetGroupArn'])['TargetHealthDescriptions']
                for vpc in response:
                    is_empty_default_vpc = fingerprint_vpc(region, vpc, ec2)
                    if not is_empty_default_vpc:
                        topology[region]['vpcs'].append(vpc)
            except botocore.exceptions.ClientError:
                rprint(f":x: [red]Client Error reported for region {region}. Most likely no VPCs exist, continuing...")

def add_network_elements_to_vpcs():
    for k, v in topology.items():
        if not k in non_region_topology_keys: # Ignore these keys, all the rest are regions
            ec2 = boto3.client('ec2',region_name=k,verify=False)
            elb = boto3.client('elbv2',region_name=k,verify=False)
            for vpc in v['vpcs']:
                rprint(f"    [yellow]Discovering network elements (subnets, route tables, etc.) for {k}/{vpc['VpcId']}...")
                subnets = [subnet for subnet in ec2.describe_subnets()['Subnets'] if subnet['VpcId'] == vpc['VpcId']]
                vpc['subnets'] = subnets
                route_tables = [rt for rt in ec2.describe_route_tables()['RouteTables'] if rt['VpcId'] == vpc['VpcId']]
                vpc['route_tables'] = route_tables
                igws = [igw for igw in ec2.describe_internet_gateways()['InternetGateways'] if igw['Attachments'][0]['VpcId'] == vpc['VpcId']]
                vpc['internet_gateways'] = igws
                natgws = [natgw for natgw in ec2.describe_nat_gateways()['NatGateways'] if natgw['VpcId'] == vpc['VpcId']]
                vpc['nat_gateways'] = natgws
                eigws = [eigw for eigw in ec2.describe_egress_only_internet_gateways()['EgressOnlyInternetGateways'] if eigw['Attachments'][0]['VpcId'] == vpc['VpcId']]
                vpc['egress_only_internet_gateways'] = eigws
                sec_grps = [sg for sg in ec2.describe_security_groups()['SecurityGroups'] if sg['VpcId'] == vpc['VpcId']]
                vpc['security_groups'] = sec_grps
                net_acls = [acl for acl in ec2.describe_network_acls()['NetworkAcls'] if acl['VpcId'] == vpc['VpcId']]
                vpc['network_acls'] = net_acls
                vpn_gateways = [gw for gw in ec2.describe_vpn_gateways()['VpnGateways'] for attch in gw['VpcAttachments'] if attch['VpcId'] == vpc['VpcId']]
                vpc['vpn_gateways'] = vpn_gateways
                for gw in vpn_gateways: # Add VPN connections to owner VPN gateways
                    gw['connections'] = [conn for conn in ec2.describe_vpn_connections()['VpnConnections'] if conn['VpnGatewayId'] == gw['VpnGatewayId']]
                    cgw_ids = [cgw['CustomerGatewayId'] for cgw in gw['connections']]
                    gw['customer_gateways'] = [cgw for cgw in ec2.describe_customer_gateways()['CustomerGateways'] if cgw['CustomerGatewayId'] in cgw_ids]
                ec2_instances = [inst for each in ec2.describe_instances()['Reservations'] for inst in each['Instances'] if "VpcId" in inst.keys() and inst['VpcId'] == vpc['VpcId']]
                ec2_groups = [grp for each in ec2.describe_instances()['Reservations'] for grp in each['Groups']]
                vpc['ec2_instances'] = ec2_instances
                vpc['ec2_groups'] = ec2_groups
                vpc['endpoints'] = [ep for ep in ec2.describe_vpc_endpoints()['VpcEndpoints'] if ep['VpcId'] == vpc['VpcId']]
                vpc['load_balancers'] = [lb for lb in elb.describe_load_balancers()['LoadBalancers'] if lb['VpcId'] == vpc['VpcId']]
                for lb in vpc['load_balancers']:
                    lb['Listeners'] = elb.describe_listeners(LoadBalancerArn=lb['LoadBalancerArn'])['Listeners']
                vpc['lb_target_groups'] = [tg for tg in elb.describe_target_groups()['TargetGroups'] if "VpcId" in tg.keys() and tg['VpcId'] == vpc['VpcId']]
                for tg in vpc['lb_target_groups']:
                    tg['HealthChecks'] = elb.describe_target_health(TargetGroupArn=tg['TargetGroupArn'])['TargetHealthDescriptions']

def add_prefix_lists_to_topology():
    for region in topology:
        if not region in non_region_topology_keys:
            rprint(f"    [yellow]Interrogating Region {region} for Prefix Lists...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            try:
                pls = [pl for pl in ec2.describe_prefix_lists()['PrefixLists']]
                topology[region]['prefix_lists'] = pls
            except botocore.exceptions.ClientError:
                rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_vpn_customer_gateways_to_topology():
    for region, v in topology.items():
        if not region in non_region_topology_keys: # Ignore these keys, all the rest are regions
            rprint(f"    [yellow]Interrogating Region {region} for Customer Gateways...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            try:
                v['customer_gateways'] = [cgw for cgw in ec2.describe_customer_gateways()['CustomerGateways']]
            except botocore.exceptions.ClientError:
                rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_vpn_tgw_connections_to_topology():
    for region, v in topology.items():
        if not region in non_region_topology_keys: # Ignore these keys, all the rest are regions
            rprint(f"    [yellow]Interrogating Region {region} for VPN Connections Attached to Transit Gateways...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            try:
                v['vpn_tgw_connections'] = [conn for conn in ec2.describe_vpn_connections()['VpnConnections'] if "TransitGatewayId" in conn.keys()]
            except botocore.exceptions.ClientError:
                rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_endpoint_services_to_topology():
    for region, v in topology.items():
        if not region in non_region_topology_keys: # Ignore these keys, all the rest are regions
            rprint(f"    [yellow]Interrogating Region {region} for Endpoint Services...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            try:
                v['endpoint_services'] = [svc for svc in ec2.describe_vpc_endpoint_services()['ServiceDetails'] if not svc['Owner'] == "amazon"]
            except botocore.exceptions.ClientError:
                rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_vpc_peering_connections_to_topology():
    pcx = [conn for conn in ec2.describe_vpc_peering_connections()['VpcPeeringConnections']]
    topology['vpc_peering_connections'] = pcx

def add_direct_connect_to_topology():
    dx = boto3.client('directconnect', verify=False)
    dcgws = [dcgw for dcgw in dx.describe_direct_connect_gateways()['directConnectGateways']]
    topology['direct_connect_gateways'] = dcgws
    for dcgw in topology['direct_connect_gateways']:
        dcgw['Attachments'] = [attch for attch in dx.describe_direct_connect_gateway_attachments(directConnectGatewayId=dcgw['directConnectGatewayId'])['directConnectGatewayAttachments']]
        dcgw['Associations'] = [assoc for assoc in dx.describe_direct_connect_gateway_associations(directConnectGatewayId=dcgw['directConnectGatewayId'])['directConnectGatewayAssociations']]
        virtual_interfaces = [vif for vif in dx.describe_virtual_interfaces()['virtualInterfaces'] if vif['directConnectGatewayId'] == dcgw['directConnectGatewayId']]
        dcgw['Connections'] = [conn for conn in dx.describe_connections()['connections'] if conn['connectionId'] in [vif['connectionId'] for vif in virtual_interfaces]]
        for conn in dcgw['Connections']:
            conn['VirtualInterfaces'] = [vif for vif in virtual_interfaces if vif['connectionId'] == conn['connectionId']]

def add_transit_gateways_to_topology():
    for region in topology:
        if not region in non_region_topology_keys: # Ignore these dictionary keys, they're not a region, all others are regions
            rprint(f"    [yellow]Interrogating Region {region} for Transit Gateways...")
            ec2 = boto3.client('ec2',region_name=region,verify=False)
            try:
                tgws = [tgw for tgw in ec2.describe_transit_gateways()['TransitGateways']]
                for tgw in tgws:
                    attachments = [attachment for attachment in ec2.describe_transit_gateway_attachments()['TransitGatewayAttachments'] if attachment['TransitGatewayId'] == tgw['TransitGatewayId']]
                    for attachment in attachments: # Loop through VPC attachments and set ApplianceModeSupport option
                        appliance_mode_support = "disable"
                        if attachment['ResourceType'] == "vpc":
                            attachment_options = ec2.describe_transit_gateway_vpc_attachments(Filters=[{'Name':attachment['TransitGatewayAttachmentId']}])['TransitGatewayVpcAttachments'][0]
                            appliance_mode_support = attachment_options['Options']['ApplianceModeSupport']
                        attachment['ApplianceModeSupport'] = appliance_mode_support
                    tgw['attachments'] = attachments
                    rts = [rt for rt in ec2.describe_transit_gateway_route_tables()['TransitGatewayRouteTables'] if rt['TransitGatewayId'] == tgw['TransitGatewayId']]
                    tgw['route_tables'] = rts
                topology[region]['transit_gateways'] = tgws
            except botocore.exceptions.ClientError as e:
                if "(UnauthorizedOperation)" in str(e):
                    rprint(f"[red]Unauthorized Operation reported while pulling Transit Gateways from {region}. Skipping...")
                else:
                    print(e)

def add_transit_gateway_routes_to_topology():
    for region in topology:
       if not region in non_region_topology_keys: # Ignore these dictionary keys, they're not a region, all others are regions
            rprint(f"    [yellow]Interrogating Region {region} for Transit Gateway Routes...")
            ec2 = boto3.client('ec2',region_name=region,verify=False) 
            tgw_routes = []
            try:
                tgw_rts = [rt for rt in ec2.describe_transit_gateway_route_tables()['TransitGatewayRouteTables']]
                for rt in tgw_rts:
                    routes = [route for route in ec2.search_transit_gateway_routes(
                        TransitGatewayRouteTableId = rt['TransitGatewayRouteTableId'],
                        Filters = [
                            {
                                "Name": "state",
                                "Values": ["active","blackhole"]
                            }
                        ]
                    )['Routes']]
                    tgw_routes.append({
                        "TransitGatewayRouteTableId": rt['TransitGatewayRouteTableId'],
                        "TransitGatewayRouteTableName": extract_name_from_aws_tags(rt['Tags']),
                        "Routes":routes
                    })
            except botocore.exceptions.ClientError as e:
                if "(UnauthorizedOperation)" in str(e):
                    rprint(f"[red]Unauthorized Operation reported while pulling Transit Gateway Route Tables from {region}. Skipping...")
                else:
                    print(e)
            topology[region]['transit_gateway_routes'] = tgw_routes

# BUILD WORD TABLE FUNCTIONS
def add_vpcs_to_word_doc():
    # Create the base table model
    vpc_model = deepcopy(word_table_models.vpc_tbl)
    # Populate the table model with data
    vpcs = [{"region":region,"vpc":vpc} for region, children in topology.items() if "vpcs" in children and children['vpcs'] for vpc in children['vpcs']]
    for rownum, vpc in enumerate(sorted(vpcs, key = lambda d : d['region']), start=1):
        this_rows_cells = []
        # Shade every other row for readability
        if not (rownum % 2) == 0:
            row_color = alternating_row_color
        else:
            row_color = None
        try: # Get VPC Name (from tag)
            vpc_name = [tag['Value'] for tag in vpc['vpc']['Tags'] if tag['Key'] == "Name"][0]
        except KeyError:
            vpc_name = ""
        except IndexError:
            vpc_name = ""
        # Get number of instances in this VPC
        inst_qty = str(len(vpc['vpc']['ec2_instances']))
        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc['region']}]})
        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc_name}]})
        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc['vpc']['CidrBlock']}]})
        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc['vpc']['VpcId']}]})
        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":inst_qty}]})
        # inject the row of cells into the table model
        vpc_model['table']['rows'].append({"cells":this_rows_cells})
    # Model has been build, now convert it to a python-docx Word table object
    table = build_table(doc_obj, vpc_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpcs}}", table)

def add_route_tables_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.rt_tbl)
                for rownum, rt in enumerate(vpc['route_tables'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get Route Table name
                        rt_name = [tag['Value'] for tag in rt['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        rt_name = ""
                    except IndexError:
                        # Object has no name
                        rt_name = ""
                    # Get number of routes
                    route_qty = len(rt['Routes'])
                    # Get number of subnet associations
                    subnet_associations = len([assoc for assoc in rt['Associations'] if "SubnetId" in assoc.keys()])
                    # Get number of edge associations
                    edge_associations = len([assoc for assoc in rt['Associations'] if "GatewayId" in assoc.keys()])
                    # Get Route Propagations
                    propagations = [x['GatewayId'] for x in rt['PropagatingVgws']]
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt['RouteTableId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(route_qty)}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(subnet_associations)}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(edge_associations)}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":propagations}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_rts}}", table)

def add_routes_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for rt in vpc['route_tables']:
                    this_parent_tbl_rows_cells = []
                    try: # Get Route Table name
                        rt_name = [tag['Value'] for tag in rt['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        rt_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        rt_name = "<unnamed>"
                    # Create the parent table row and cells
                    this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} / RT: {rt_name} ({rt['RouteTableId']})"}]})
                    # inject the row of cells into the table model
                    parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                    # Build the child table
                    child_model = deepcopy(word_table_models.rt_routes_tbl)
                    for rownum, route in enumerate(rt['Routes'], start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        # Get Destination
                        if "DestinationCidrBlock" in route.keys():
                            destination = route['DestinationCidrBlock']
                        elif "DestinationPrefixListId" in route.keys():
                            destination = route['DestinationPrefixListId']
                        else:
                            destination = "Unknown Destination Type"
                        # Get Destination Gateway
                        if "GatewayId" in route.keys():
                            gateway = route['GatewayId']
                        elif "TransitGatewayId" in route.keys():
                            gateway = route['TransitGatewayId']
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":destination}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":gateway}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":route['Origin']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_rt_routes}}", table)

def add_prefix_lists_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if isinstance(attributes, dict) and "prefix_lists" in attributes.keys():
            # Populate the table model with data
            if not attributes['prefix_lists']:
                pass
            else:
                this_parent_tbl_rows_cells = []
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.prefix_list_tbl)
                for rownum, pl in enumerate(attributes['prefix_lists'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":pl['PrefixListName']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":pl['PrefixListId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":pl['Cidrs']}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no Prefix Lists at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Prefix Lists Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_prefix_lists}}", table)

def add_subnets_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name == ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.subnet_tbl)
                for rownum, subnet in enumerate(sorted(vpc['subnets'], key = lambda d : d['AvailabilityZone']), start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get the Subnet Name
                        subnet_name = [tag['Value'] for tag in subnet['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError as e:
                        # Object has no name
                        subnet_name = ""
                    except IndexError:
                        subnet_name = ""
                    # Get Route Table
                    try:
                        route_table = [rt['RouteTableId'] for rt in vpc['route_tables'] for assoc in rt['Associations'] if "SubnetId" in assoc.keys() and assoc['SubnetId'] == subnet['SubnetId']][0]
                    except IndexError:
                        route_table = ""
                    # Get Network ACLs
                    net_acls = [assoc['NetworkAclId'] for acl in vpc['network_acls'] for assoc in acl['Associations'] if assoc['SubnetId'] == subnet['SubnetId']]
                    # Get number of instances in this subnet
                    inst_qty = str(len([inst['SubnetId'] for inst in vpc['ec2_instances'] for intf in inst['NetworkInterfaces'] if intf['SubnetId'] == subnet['SubnetId']]))
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnet['CidrBlock']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnet_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnet['AvailabilityZone']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":route_table}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":net_acls}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnet['SubnetId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":inst_qty}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_subnets}}", table)

def add_network_acls_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.netacls_tbl)
                for rownum, acl in enumerate(vpc['network_acls'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get ACL Table name
                        acl_name = [tag['Value'] for tag in acl['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        acl_name = ""
                    except IndexError:
                        # Object has no name
                        acl_name = ""
                    # Get Subnets associated with ACL
                    subnet_associations = [assoc['SubnetId'] for assoc in acl['Associations']]
                    # Get IsDefault status
                    is_default = "yes" if acl['IsDefault'] else "no"
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":acl_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":acl['NetworkAclId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":is_default}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnet_associations}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_netacls}}", table)

def add_netacl_inbound_entries_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for acl in vpc['network_acls']:
                    this_parent_tbl_rows_cells = []
                    try: # Get ACL name
                        acl_name = [tag['Value'] for tag in acl['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        acl_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        acl_name = "<unnamed>"
                    # Create the parent table row and cells
                    this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} / ACL: {acl_name} ({acl['NetworkAclId']})"}]})
                    # inject the row of cells into the table model
                    parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                    # Build the child table
                    inbound_entries = [entry for entry in acl['Entries'] if not entry['Egress']]
                    child_model = deepcopy(word_table_models.netacl_in_entries_tbl)
                    for rownum, entry in enumerate(sorted(inbound_entries, key = lambda d : d['RuleNumber']), start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        # Get Port Range
                        try:
                            if entry['PortRange']['From'] == entry['PortRange']['To']:
                                port_range = str(entry['PortRange']['From'])
                            else:
                                port_range = f"{entry['PortRange']['From']}-{entry['PortRange']['To']}"
                        except KeyError:
                            port_range = "All"
                        try: # Get CIDR Block
                            cidr_block = entry['CidrBlock']
                        except KeyError:
                            cidr_block = ""
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(entry['RuleNumber'])}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":aws_protocol_map[entry['Protocol']]}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":port_range}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cidr_block}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['RuleAction']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_netacl_in_entries}}", table)

def add_netacl_outbound_entries_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for acl in vpc['network_acls']:
                    this_parent_tbl_rows_cells = []
                    try: # Get ACL name
                        acl_name = [tag['Value'] for tag in acl['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        acl_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        acl_name = "<unnamed>"
                    # Create the parent table row and cells
                    this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} / ACL: {acl_name} ({acl['NetworkAclId']})"}]})
                    # inject the row of cells into the table model
                    parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                    # Build the child table
                    outbound_entries = [entry for entry in acl['Entries'] if entry['Egress']]
                    child_model = deepcopy(word_table_models.netacl_in_entries_tbl)
                    for rownum, entry in enumerate(sorted(outbound_entries, key = lambda d : d['RuleNumber']), start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        # Get Port Range
                        try:
                            if entry['PortRange']['From'] == entry['PortRange']['To']:
                                port_range = str(entry['PortRange']['From'])
                            else:
                                port_range = f"{entry['PortRange']['From']}-{entry['PortRange']['To']}"
                        except KeyError:
                            port_range = "All"
                        try: # Get CIDR Block
                            cidr_block = entry['CidrBlock']
                        except KeyError:
                            cidr_block = ""
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(entry['RuleNumber'])}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":aws_protocol_map[entry['Protocol']]}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":port_range}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cidr_block}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['RuleAction']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_netacl_out_entries}}", table)

def add_security_groups_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.sec_grps_tbl)
                for rownum, sg in enumerate(vpc['security_groups'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get SG name
                        sg_name = [tag['Value'] for tag in sg['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        sg_name = ""
                    except IndexError:
                        # Object has no name
                        sg_name = ""
                    # Get Rule Counts
                    ingress_rule_count = len(sg['IpPermissions'])
                    egress_rule_count = len(sg['IpPermissionsEgress'])
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":sg_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":sg['GroupName']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":sg['GroupId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":sg['Description']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":f"{str(ingress_rule_count)}/{str(egress_rule_count)}"}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_sgs}}", table)

def add_sg_inbound_entries_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for sg in vpc['security_groups']:
                    this_parent_tbl_rows_cells = []
                    try: # Get SG name
                        sg_name = [tag['Value'] for tag in sg['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        sg_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        sg_name = "<unnamed>"
                    # Create the parent table row and cells
                    this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} / SG: {sg_name} ({sg['GroupId']})"}]})
                    # inject the row of cells into the table model
                    parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                    # Build the child table
                    inbound_entries = [entry for entry in sg['IpPermissions']]
                    # An entry can have multiple sources so we need to extract them all
                    extracted_entries = []
                    for entry in inbound_entries:
                        # Get Port Range
                        try:
                            if str(entry['FromPort']) == "-1":
                                port_range = "All"
                            if entry['FromPort'] == entry['ToPort']:
                                port_range = str(entry['FromPort'])
                            else:
                                port_range = f"{entry['FromPort']}-{entry['ToPort']}"
                        except KeyError:
                            port_range = "All"
                        # Transpose IP Protocol
                        protocol = "All" if str(entry['IpProtocol']) == "-1" else entry['IpProtocol']
                        # Build source and description
                        ip_sources = []
                        for source in entry['IpRanges']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['CidrIp'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            ip_sources.append(this_entry)
                        ipv6_sources = []
                        for source in entry['Ipv6Ranges']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['CidrIpv6'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            ipv6_sources.append(this_entry)
                        prefix_sources = []
                        for source in entry['PrefixListIds']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['PrefixListId'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            prefix_sources.append(this_entry)
                        sg_sources = []
                        for source in entry['UserIdGroupPairs']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['GroupId'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            sg_sources.append(this_entry)
                        sources = ip_sources + ipv6_sources + prefix_sources + sg_sources
                        for each in sources:
                            extracted_entries.append(each)
                    child_model = deepcopy(word_table_models.sec_grp_in_entries_tbl)
                    for rownum, entry in enumerate(extracted_entries, start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['protocol']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['port_range']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['source']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['description']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_sg_in_entries}}", table)

def add_sg_outbound_entries_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for sg in vpc['security_groups']:
                    this_parent_tbl_rows_cells = []
                    try: # Get SG name
                        sg_name = [tag['Value'] for tag in sg['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        sg_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        sg_name = "<unnamed>"
                    # Create the parent table row and cells
                    this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} / SG: {sg_name} ({sg['GroupId']})"}]})
                    # inject the row of cells into the table model
                    parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                    # Build the child table
                    inbound_entries = [entry for entry in sg['IpPermissionsEgress']]
                    # An entry can have multiple sources so we need to extract them all
                    extracted_entries = []
                    for entry in inbound_entries:
                        # Get Port Range
                        try:
                            if str(entry['FromPort']) == "-1":
                                port_range = "All"
                            if entry['FromPort'] == entry['ToPort']:
                                port_range = str(entry['FromPort'])
                            else:
                                port_range = f"{entry['FromPort']}-{entry['ToPort']}"
                        except KeyError:
                            port_range = "All"
                        # Transpose IP Protocol
                        protocol = "All" if str(entry['IpProtocol']) == "-1" else entry['IpProtocol']
                        # Build source and description
                        ip_sources = []
                        for source in entry['IpRanges']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['CidrIp'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            ip_sources.append(this_entry)
                        ipv6_sources = []
                        for source in entry['Ipv6Ranges']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['CidrIpv6'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            ipv6_sources.append(this_entry)
                        prefix_sources = []
                        for source in entry['PrefixListIds']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['PrefixListId'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            prefix_sources.append(this_entry)
                        sg_sources = []
                        for source in entry['UserIdGroupPairs']:
                            this_entry = {
                                "protocol": protocol,
                                "port_range": port_range,
                                "source": source['GroupId'],
                                "description": "" if not "Description" in source.keys() else source['Description']
                            }
                            sg_sources.append(this_entry)
                        sources = ip_sources + ipv6_sources + prefix_sources + sg_sources
                        for each in sources:
                            extracted_entries.append(each)
                    child_model = deepcopy(word_table_models.sec_grp_in_entries_tbl)
                    for rownum, entry in enumerate(extracted_entries, start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['protocol']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['port_range']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['source']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":entry['description']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_sg_out_entries}}", table)

def add_internet_gateways_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            this_parent_tbl_rows_cells = []
            # Create the parent table row and cells
            this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
            # inject the row of cells into the table model
            parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
            # Build the child table
            child_model = deepcopy(word_table_models.igw_tbl)
            for rownum, vpc in enumerate(vpcs, start=1):
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for igw in vpc['internet_gateways']:
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get IGW name
                        igw_name = [tag['Value'] for tag in igw['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        igw_name = ""
                    except IndexError:
                        # Object has no name
                        igw_name = ""
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":igw_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":igw['InternetGatewayId']}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
            # Add the child table to the parent table
            parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_igws}}", table)

def add_egress_only_internet_gateways_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            this_parent_tbl_rows_cells = []
            # Create the parent table row and cells
            this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
            # inject the row of cells into the table model
            parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
            # Build the child table
            child_model = deepcopy(word_table_models.eigw_tbl)
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for rownum, igw in enumerate(vpc['egress_only_internet_gateways'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get IGW name
                        igw_name = [tag['Value'] for tag in igw['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        igw_name = ""
                    except IndexError:
                        # Object has no name
                        igw_name = ""
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":igw_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":igw['EgressOnlyInternetGatewayId']}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
            # Add the child table to the parent table
            parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_eigws}}", table)

def add_nat_gateways_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            this_parent_tbl_rows_cells = []
            # Create the parent table row and cells
            this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
            # inject the row of cells into the table model
            parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
            # Build the child table
            child_model = deepcopy(word_table_models.ngw_tbl)
            for vpc in vpcs:
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                for rownum, ngw in enumerate(vpc['nat_gateways'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get IGW name
                        ngw_name = [tag['Value'] for tag in ngw['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        ngw_name = ""
                    except IndexError:
                        # Object has no name
                        ngw_name = ""
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vpc_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ngw_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ngw['NatGatewayId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ngw['SubnetId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ngw['ConnectivityType']}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the Addresses header
                    child_model['table']['rows'].append({"cells":[{"background": table_header_color,"paragraphs": [{"style": "regularbold", "text": "ADDRESSES"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                    # Build the address table
                    address_model = deepcopy(word_table_models.ngw_address_tbl)
                    for rownum2, address in enumerate(ngw['NatGatewayAddresses']):
                        address_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum2 % 2) == 0:
                            address_row_color = alternating_row_color
                        else:
                            address_row_color = None
                        # Get Public IP
                        public_ip = address['PublicIp'] if any("PublicIp" in key for key in address) else ""
                        # Convert IsPrimary key from bool to string
                        is_primary = "True" if address['IsPrimary'] else "False"
                        # Build word table rows & cells
                        address_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":public_ip}]})
                        address_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":address['PrivateIp']}]})
                        address_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":address['NetworkInterfaceId']}]})
                        address_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":is_primary}]})
                        # inject the row of cells into the table model
                        address_model['table']['rows'].append({"cells":address_rows_cells})
                    # Add the address table to the child table
                    child_model['table']['rows'].append({"cells":[address_model,{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
            # Add the child table to the parent table
            parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_ngws}}", table)

def add_endpoint_services_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if not region in non_region_topology_keys and "endpoint_services" in attributes.keys(): # Ignore these dictionary keys, they are not a region, and only run if endpoint services exist
            # Create Table title (Region)
            parent_model['table']['rows'].append({"cells": [{"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]}]})
            for rownum, epsvc in enumerate(sorted(attributes['endpoint_services'], key = lambda d : d['ServiceType'][0]['ServiceType']), start=1):
                if rownum > 1: # Inject a space between Endpoint Services tables
                    parent_model['table']['rows'].append({"cells": [{"paragraphs":[{"style":"No Spacing","text":""}]}]})
                # Create Child table and populate header values
                child_model = deepcopy(word_table_models.endpoint_services_tbl)
                ep_svc_name = extract_name_from_aws_tags(epsvc['Tags'])
                if len(epsvc['ServiceType']) > 1: # This script assumes only a single service type, but Amazon returns a list (so there could be more). Warn if there are more.
                    rprint("\t[orange]WARNING: This script assumes a single service type but multiple detected. Data could be missing, please let the script author know about this condition.")
                child_model['table']['rows'][0]['cells'][0]['paragraphs'].append({"style":"regularbold","text":f"ENDPOINT SERVICE NAME: {epsvc['ServiceName']}"})
                child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":ep_svc_name})
                child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":epsvc['ServiceId']})
                child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":epsvc['ServiceType'][0]['ServiceType']})
                child_model['table']['rows'][2]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":epsvc['AvailabilityZones']})
                # Add child model to parent table model
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_endpoint_services}}", table)

def add_endpoints_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name == ""
                vpc_label = vpc_name if not vpc_name == "" else vpc['VpcId']
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_label}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                if not vpc['endpoints']:
                    parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":"No Endpoints present"}]}]})
                else:
                    # Build the child table
                    child_model = deepcopy(word_table_models.endpoint_tbl)
                    for rownum, ep in enumerate(sorted(vpc['endpoints'], key = lambda d : d['VpcEndpointType']), start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        try: # Get the Subnet Name
                            ep_name = [tag['Value'] for tag in ep['Tags'] if tag['Key'] == "Name"][0]
                        except KeyError as e:
                            # Object has no name
                            ep_name = ""
                        except IndexError:
                            ep_name = ""
                        # Build Subnet ID list and cross-reference Subnet Names
                        ep_subnets = [f"{subnet}({get_subnet_name_by_id(subnet,vpc)})" for subnet in ep['SubnetIds']]
                        # ep_subnets = [f"{subnet}({extract_name_from_aws_tags(subnet2['Tags'])})" for subnet in ep['SubnetIds'] for subnet2 in vpc['subnets'] if subnet == subnet2['SubnetId']]
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep_name}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep['VpcEndpointId']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep['VpcEndpointType']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep['NetworkInterfaceIds']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep_subnets}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ep['ServiceName']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Endpoints Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_endpoints}}", table)

def add_vpc_peerings_to_word_doc():
    # Create the base table model
    model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    if not topology['vpc_peering_connections']:
        model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPC Peerings"}]},{"merge":None},{"merge":None},{"merge":None}]})
    else:
        for rownum, pcx in enumerate(topology['vpc_peering_connections']):
            if rownum > 0: # Inject an empty row to space the data
                model['table']['rows'].append({"cells":[]})
            this_rows_requester_cells = []
            this_rows_accepter_cells = []
            try: # Get VPC Peering name
                pcx_name = [tag['Value'] for tag in pcx['Tags'] if tag['Key'] == "Name"][0]
            except KeyError:
                # Object has no name
                pcx_name = ""
            except IndexError:
                # Object has no name
                pcx_name = ""
            # Create child table model & populate header row with data
            child_model = deepcopy(word_table_models.vpc_peering_requester_tbl)
            child_model['table']['rows'][0]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":pcx_name})
            child_model['table']['rows'][0]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":pcx['VpcPeeringConnectionId']})
            # Populate requester data row with data
            this_rows_requester_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['RequesterVpcInfo']['Region']}]})
            this_rows_requester_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['RequesterVpcInfo']['VpcId']}]})
            this_rows_requester_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['RequesterVpcInfo']['CidrBlock']}]})
            this_rows_requester_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['RequesterVpcInfo']['OwnerId']}]})
            # inject the requester row of cells into the child table model
            child_model['table']['rows'].append({"cells":this_rows_requester_cells})
            # Populate accepter data row with data
            this_rows_accepter_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['AccepterVpcInfo']['Region']}]})
            this_rows_accepter_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['AccepterVpcInfo']['VpcId']}]})
            this_rows_accepter_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['AccepterVpcInfo']['CidrBlock']}]})
            this_rows_accepter_cells.append({"paragraphs":[{"style":"No Spacing","text":pcx['AccepterVpcInfo']['OwnerId']}]})
            # inject the accepter header cells and data row of cells into the child table model
            child_model['table']['rows'].append(word_table_models.vpc_peering_accepter_tbl_header)
            child_model['table']['rows'].append({"cells":this_rows_accepter_cells})
            # Add child model to parent table model
            model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    table = build_table(doc_obj, model)
    replace_placeholder_with_table(doc_obj, "{{py_pcx}}", table)

def add_transit_gateways_to_word_doc():
    # Create the parent table model
    model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if not region in non_region_topology_keys and attributes['transit_gateways']: # Ignore these dictionary keys, they are not a region, also don't run if there are no transit gateways in the region
            # Create Table title (Region)
            model['table']['rows'].append({"cells": [{"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]}]})
            for rownum, tgw in enumerate(attributes['transit_gateways']):
                if rownum > 0: # Inject an empty row to space the data
                    model['table']['rows'].append({"cells":[]})
                try: # Get TGW name
                    tgw_name = [tag['Value'] for tag in tgw['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    tgw_name = ""
                except IndexError:
                    # Object has no name
                    tgw_name = ""
                # Create child table model & populate header rows with data
                child_model = deepcopy(word_table_models.tgw_tbl)
                child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":tgw_name})
                child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":tgw['TransitGatewayId']})
                child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":str(tgw['Options']['AmazonSideAsn'])})
                child_model['table']['rows'][2]['cells'][3]['paragraphs'].append({"style":"No Spacing" if tgw['OwnerId'] == topology['account']['id'] else "redtext","text":tgw['OwnerId']})
                # Populate child table model with spacer and attachment header
                child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style":"regularbold","text":"ATTACHMENTS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                child_model['table']['rows'].append(deepcopy(word_table_models.tgw_attachment_tbl_header))
                # Populate child table with attachments
                for rownum2, attch in enumerate(tgw['attachments'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum2 % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get TGW Attachment name
                        attch_name = [tag['Value'] for tag in attch['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        attch_name = ""
                    except IndexError:
                        # Object has no name
                        attch_name = ""
                    try: # Get TGW Route Table ID
                        rt_id = attch['Association']['TransitGatewayRouteTableId']
                    except KeyError:
                        rt_id = ""
                    # Set Resource Type (if VPC, check appliance mode and append to output)
                    resource_type = attch['ResourceType']
                    if resource_type == "vpc" and attch['ApplianceModeSupport'] == "enable":
                        resource_type = f"{attch['ResourceType']} (Appliance Mode Enabled)"
                    # Add data to row/cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch['TransitGatewayAttachmentId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch['ResourceType']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch['ResourceId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt_id}]})
                    # add attachment data row to child table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Populate child table model with spacer and route table header
                child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs":[{"style":"regularbold","text":"ROUTE TABLES"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                child_model['table']['rows'].append(deepcopy(word_table_models.tgw_rt_tbl_header))
                for rownum2, rt in enumerate(tgw['route_tables'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum2 % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get TGW RT name
                        rt_name = [tag['Value'] for tag in rt['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        rt_name = "<unnamed>"
                    except IndexError:
                        # Object has no name
                        rt_name = "<unnamed>"
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt_name}]})
                    this_rows_cells.append({"merge":None})
                    this_rows_cells.append({"merge":None})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt['TransitGatewayRouteTableId']}]})
                    this_rows_cells.append({"merge":None})
                    # add route table header row to child table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add child model to parent table model
                model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    table = build_table(doc_obj, model)
    replace_placeholder_with_table(doc_obj, "{{py_tgws}}", table)

def add_transit_gateway_routes_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if not region in non_region_topology_keys: # Ignore these dictionary keys, they are not a region
            if not attributes['transit_gateway_routes']:
                pass
            else:
                for rt in attributes['transit_gateway_routes']:
                    parent_model['table']['rows'].append(
                        {"cells":[{"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / RT: {rt['TransitGatewayRouteTableName']} ({rt['TransitGatewayRouteTableId']})"}]}]}
                    )
                    # Build the child table
                    child_model = deepcopy(word_table_models.tgw_routes_tbl)
                    for rownum, route in enumerate(sorted(rt['Routes'], key = lambda d : d['Type']), start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        try: # Get Resource Type
                            resource_type = route['TransitGatewayAttachments'][0]['ResourceType']
                        except KeyError:
                            resource_type = "-"
                        try: # Get Resource ID
                            resource_id = route['TransitGatewayAttachments'][0]['ResourceId']
                        except KeyError:
                            resource_id = "-"
                        try: # Get Attachment ID
                            attachment_id = route['TransitGatewayAttachments'][0]['TransitGatewayAttachmentId']
                        except KeyError:
                            attachment_id = "-"
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":route['DestinationCidrBlock']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":resource_type}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":resource_id}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attachment_id}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":route['Type']}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":route['State']}]})
                        # inject the row of cells into the table model
                        child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Transit Gateway Routes Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_tgw_routes}}", table)

def add_vpn_customer_gateways_to_word():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if isinstance(attributes, dict) and "customer_gateways" in attributes.keys():
            # Populate the table model with data
            if not attributes['customer_gateways']:
                pass
            else:
                this_parent_tbl_rows_cells = []
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                child_model = deepcopy(word_table_models.vpn_cgw_tbl)
                for rownum, cgw in enumerate(attributes['customer_gateways'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    try: # Get CGW name
                        cgw_name = [tag['Value'] for tag in cgw['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        cgw_name = ""
                    except IndexError:
                        # Object has no name
                        cgw_name = ""
                    try: # Get CGW Device name
                        cgw_dev_name = cgw['DeviceName']
                    except KeyError:
                        # Object has no name
                        cgw_dev_name = ""
                    except IndexError:
                        # Object has no name
                        cgw_dev_name = ""
                    # Build word table rows & cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw_dev_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['IpAddress']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['BgpAsn']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['CustomerGatewayId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['Type']}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no Prefix Lists at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Customer Gateways attached to transit gateways present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_vpn_cgws}}", table)

def add_vpn_tgw_connections_to_word():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if isinstance(attributes, dict) and "vpn_tgw_connections" in attributes.keys():
            # Populate the table model with data
            if not attributes['vpn_tgw_connections']:
                pass
            else:
                this_parent_tbl_rows_cells = []
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                for rownum, conn in enumerate(attributes['vpn_tgw_connections'], start=1):
                    child_model = deepcopy(word_table_models.vpn_tgw_conn_tbl)
                    if rownum > 1: # Add a line break between connections for readability
                        parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":""}]}]})
                    try: # Get Connection name
                        conn_name = [tag['Value'] for tag in conn['Tags'] if tag['Key'] == "Name"][0]
                    except KeyError:
                        # Object has no name
                        conn_name = ""
                    except IndexError:
                        # Object has no name
                        conn_name = ""
                    # Build word table rows & cells
                    child_model['table']['rows'][0]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":conn_name})
                    child_model['table']['rows'][0]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":conn['VpnConnectionId']})
                    child_model['table']['rows'][0]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":conn['TransitGatewayId']})
                    child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":conn['CustomerGatewayId']})
                    child_model['table']['rows'][1]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":conn['Type']})
                    next_row = []
                    next_row.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['LocalIpv4NetworkCidr']}]})
                    next_row.append({"merge":None})
                    next_row.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['RemoteIpv4NetworkCidr']}]})
                    next_row.append({"merge":None})
                    next_row.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['OutsideIpAddressType']}]})
                    next_row.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['TunnelInsideIpVersion']}]})
                    child_model['table']['rows'].append({"cells":next_row})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
                    conn_label = conn_name if not conn_name == "" else conn['VpnConnectionId']
                    child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs":[{"style":"regularbold","text":f"{conn_label} VPN CONNECTION TUNNELS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                    child_model['table']['rows'].append(word_table_models.vgw_conn_tunnel_tbl_header)
                    for rownum2, tun in enumerate(conn['Options']['TunnelOptions'], start=1):
                        this_rows_cells = []
                        # Shade every other row for readability
                        if not (rownum2 % 2) == 0:
                            row_color = alternating_row_color
                        else:
                            row_color = None
                        # Get this tunnels IPSec status and tunnel status
                        ipsec_status = [status['StatusMessage'] for status in conn['VgwTelemetry'] if status['OutsideIpAddress'] == tun['OutsideIpAddress']][0]
                        tun_status = [status['Status'] for status in conn['VgwTelemetry'] if status['OutsideIpAddress'] == tun['OutsideIpAddress']][0]
                        # Add connection rows to table
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun['OutsideIpAddress']}]})
                        this_rows_cells.append({"merge":None})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun['TunnelInsideCidr']}]})
                        this_rows_cells.append({"merge":None})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ipsec_status}]})
                        this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun_status}]})
                        # inject cells into the child table row
                        child_model['table']['rows'].append({"cells":this_rows_cells})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no Prefix Lists at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPN Transit Gateway Connections present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_vpn_s2s}}", table)

def add_vpn_gateways_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                if not vpc['vpn_gateways']:
                    parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style": "No Spacing", "text": "No VPN Gateways configured"}]}]})
                else:
                    # Build the child table
                    child_model = deepcopy(word_table_models.vgw_tbl)
                    for rownum, gw in enumerate(vpc['vpn_gateways'], start=1):
                        if rownum > 1: # Inject an empty row to space the data
                            child_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": ""}]},{"merge":None},{"merge":None},{"merge":None}]})
                        try: # Get VGW name
                            gw_name = [tag['Value'] for tag in gw['Tags'] if tag['Key'] == "Name"][0]
                        except KeyError:
                            # Object has no name
                            gw_name = ""
                        except IndexError:
                            # Object has no name
                            gw_name = ""
                        # Build word table rows & cells
                        child_model['table']['rows'][0]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":gw_name})
                        child_model['table']['rows'][0]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":gw['VpnGatewayId']})
                        child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":str(gw['AmazonSideAsn'])})
                        child_model['table']['rows'][1]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":gw['Type']})
                        child_model['table']['rows'][2]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":[vpc['VpcId'] for vpc in gw['VpcAttachments']]})
                        # Insert the Customer Gateway Header into the Child Table
                        child_model['table']['rows'].append(word_table_models.vgw_cgw_tbl_header)
                        # Add associated Customer Gateways
                        if not gw['customer_gateways']:
                            child_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Customer Gateways"}]},{"merge":None},{"merge":None},{"merge":None}]})
                        else:
                            for rownum2, cgw in enumerate(gw['customer_gateways'], start=1):
                                this_rows_cells = []
                                # Shade every other row for readability
                                if not (rownum2 % 2) == 0:
                                    row_color = alternating_row_color
                                else:
                                    row_color = None
                                try: # Get CGW name
                                    cgw_name = [tag['Value'] for tag in cgw['Tags'] if tag['Key'] == "Name"][0]
                                except KeyError:
                                    # Object has no name
                                    cgw_name = ""
                                except IndexError:
                                    # Object has no name
                                    cgw_name = ""
                                try: # Get CGW Device name
                                    cgw_dev_name = cgw['DeviceName']
                                except KeyError:
                                    # Object has no name
                                    cgw_dev_name = ""
                                except IndexError:
                                    # Object has no name
                                    cgw_dev_name = ""
                                # Build word table rows & cells
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw_name}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['CustomerGatewayId']}]})
                                this_rows_cells.append({"merge":None})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw_dev_name}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":cgw['IpAddress']}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(cgw['BgpAsn'])}]})
                                # inject the row of cells into the table model
                                child_model['table']['rows'].append({"cells":this_rows_cells})
                        # Create new section in child table for connections
                        if not gw['connections']:
                            child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style":"regularbold","text":"VPN CONNECTION"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                            child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style":"No Spacing","text":"no vpn connections present"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                        else:
                            for conn in gw['connections']:
                                child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style":"regularbold","text":f"{conn['CustomerGatewayId']} VPN CONNECTIONS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                                child_model['table']['rows'].append(word_table_models.vgw_conn_tbl_header)
                                this_rows_cells = []
                                try: # Get Connection name
                                    conn_name = [tag['Value'] for tag in conn['Tags'] if tag['Key'] == "Name"][0]
                                except KeyError:
                                    # Object has no name
                                    conn_name = ""
                                except IndexError:
                                    # Object has no name
                                    conn_name = ""
                                # Add connection rows to table
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn_name}]})
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn['CustomerGatewayId']}]})
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['LocalIpv4NetworkCidr']}]})
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['RemoteIpv4NetworkCidr']}]})
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['OutsideIpAddressType']}]})
                                this_rows_cells.append({"paragraphs":[{"style":"No Spacing","text":conn['Options']['TunnelInsideIpVersion']}]})
                                # inject cells into the child table row
                                child_model['table']['rows'].append({"cells":this_rows_cells})
                                # Create new section in child table for connection tunnels
                                conn_label = conn_name if not conn_name == "" else conn['VpnConnectionId']
                                child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs":[{"style":"regularbold","text":f"{conn_label} VPN CONNECTION TUNNELS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                                child_model['table']['rows'].append(word_table_models.vgw_conn_tunnel_tbl_header)
                                for rownum2, tun in enumerate(conn['Options']['TunnelOptions'], start=1):
                                    this_rows_cells = []
                                    # Shade every other row for readability
                                    if not (rownum2 % 2) == 0:
                                        row_color = alternating_row_color
                                    else:
                                        row_color = None
                                    # Get this tunnels IPSec status and tunnel status
                                    ipsec_status = [status['StatusMessage'] for status in conn['VgwTelemetry'] if status['OutsideIpAddress'] == tun['OutsideIpAddress']][0]
                                    tun_status = [status['Status'] for status in conn['VgwTelemetry'] if status['OutsideIpAddress'] == tun['OutsideIpAddress']][0]
                                    # Add connection rows to table
                                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun['OutsideIpAddress']}]})
                                    this_rows_cells.append({"merge":None})
                                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun['TunnelInsideCidr']}]})
                                    this_rows_cells.append({"merge":None})
                                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":ipsec_status}]})
                                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tun_status}]})
                                    # inject cells into the child table row
                                    child_model['table']['rows'].append({"cells":this_rows_cells})
                    # Add the child table to the parent table
                    parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_vpn_vpgs}}", table)

def add_direct_connect_gateways_to_word_doc():
    # Create the parent table model
    model = deepcopy(word_table_models.parent_tbl)
    if not topology['direct_connect_gateways']:
        model['table']['rows'].append({"cells":[{"paragraphs":[{"style": "No Spacing","text":"No Direct Connect Gateways present"}]}]})
    else:
        # Populate the table model with data
        for i, gw in enumerate(topology['direct_connect_gateways']):
            if i > 0: # Inject a space between gateways
                model['table']['rows'].append({"cells":[{"paragraphs":[{"style": "No Spacing","text":""}]}]})
            # Create child table model & populate header rows with data
            child_model = deepcopy(word_table_models.dcgw_tbl)
            child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":gw['directConnectGatewayName']})
            child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":gw['directConnectGatewayId']})
            child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":str(gw['amazonSideAsn'])})
            # Add connections to child model
            for conn in gw['Connections']:
                # Insert Connections Spacer
                child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style": "regularbold","text":"CONNECTION"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                conn_header = deepcopy(word_table_models.dcgw_conn_rows)
                jumbo_frame_capable = "Yes" if conn['jumboFrameCapable'] else "No"
                macsec_capable = "Yes" if conn['macSecCapable'] else "No"
                conn_header[0]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":conn['connectionName']})
                conn_header[0]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":conn['connectionId']})
                conn_header[0]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":conn['region']})
                conn_header[1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":conn['location']})
                conn_header[1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":conn['partnerName']})
                conn_header[1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":conn['bandwidth']})
                conn_header[2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":jumbo_frame_capable})
                conn_header[2]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":macsec_capable})
                conn_header[2]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":conn['portEncryptionStatus']})
                child_model['table']['rows'] += conn_header
                # Insert Virtual Interfaces Spacer & header
                child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs":[{"style": "regularbold","text":f"CONNECTION '{conn['connectionName']}' VIRTUAL INTERFACES"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                child_model['table']['rows'].append(deepcopy(word_table_models.dcgw_vif_header))
                for rownum, vif in enumerate(conn['VirtualInterfaces'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    # Add data to row cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vif['virtualInterfaceName']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vif['virtualInterfaceType']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vif['virtualInterfaceId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":vif['amazonAddress']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(vif['mtu'])}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":[f"{peer['customerAddress']}:{peer['bgpStatus']}" for peer in vif['bgpPeers']]}]})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
            # Add child model to parent table model
            model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    table = build_table(doc_obj, model)
    replace_placeholder_with_table(doc_obj, "{{py_dcgws}}", table)

def add_load_balancers_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name == ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                if not vpc['load_balancers']:
                    parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Load Balancers Present"}]}]})
                else:
                    for rownum, lb in enumerate(sorted(vpc['load_balancers'], key = lambda d : d['Type']), start=1):
                        # Build the child table
                        child_model = deepcopy(word_table_models.load_balancer_tbl)
                        if rownum > 1: # Inject a space between load balancer tables
                            parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":""}]}]})
                        child_model['table']['rows'][0]['cells'][0]['paragraphs'].append({"style": "regularbold", "text": f"LOAD BALANCER ARN: {lb['LoadBalancerArn']}"})
                        child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":lb['LoadBalancerName']})
                        child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":lb['Type']})
                        child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":lb['State']['Code']})
                        # Populate Network Mappings
                        for rownum2, az in enumerate(lb['AvailabilityZones'], start=1):
                            this_rows_cells = []
                            # Shade every other row for readability
                            if not (rownum2 % 2) == 0:
                                row_color = alternating_row_color
                            else:
                                row_color = None
                            try: # Get Load Balancer Addresses
                                lb_addresses = "<none>" if not az['LoadBalancerAddresses'] else az['LoadBalancerAddresses']
                            except KeyError:
                                lb_addresses = "<none>"
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":az['ZoneName']}]})
                            this_rows_cells.append({"merge":None})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":f"{az['SubnetId']}({get_subnet_name_by_id(az['SubnetId'], vpc)})"}]})
                            this_rows_cells.append({"merge":None})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":lb_addresses}]})
                            this_rows_cells.append({"merge":None})
                            # inject the row of cells into the table model
                            child_model['table']['rows'].append({"cells":this_rows_cells})
                        # Insert Listener Spacer, listener header, and populate Listeners
                        child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs": [{"style": "regularbold", "text": "LISTENERS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                        child_model['table']['rows'].append(deepcopy(word_table_models.load_balancer_listener_header))
                        for rownum2, listener in enumerate(lb['Listeners'], start=1):
                            this_rows_cells = []
                            # Shade every other row for readability
                            if not (rownum2 % 2) == 0:
                                row_color = alternating_row_color
                            else:
                                row_color = None
                            # Warn if more than 1 default action or target group
                            if len(listener['DefaultActions']) > 1:
                                rprint("    [orange]WARNING: Multiple Default Actions detected in load balancer object but script only expects one. Data may be missing, please notify script author.")
                            try: # Derive Target Group from ARN
                                tg_names = [f"{tg['TargetGroupArn'].split('/')[1]}(weight:{tg['Weight']})" for tg in listener['DefaultActions'][0]['ForwardConfig']['TargetGroups']]
                            except KeyError:
                                tg_names = "---"
                            try: # Get Protocol (Gateway Load Balancers don't have a default protocol and port in the listener, so if none exists we look into the target group)
                                listener_protocol = listener['Protocol']
                            except KeyError:
                                tg = [tg for tg in vpc['lb_target_groups'] if tg['TargetGroupArn'] == listener['DefaultActions'][0]['TargetGroupArn']][0]
                                listener_protocol = tg['Protocol']
                            try: # Get Port
                                listener_port = listener['Port']
                            except KeyError:
                                tg = [tg for tg in vpc['lb_target_groups'] if tg['TargetGroupArn'] == listener['DefaultActions'][0]['TargetGroupArn']][0]
                                listener_port = tg['Port']
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":f"{listener_protocol}:{listener_port}"}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":tg_names}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":listener['ListenerArn']}]})
                            this_rows_cells.append({"merge":None})
                            this_rows_cells.append({"merge":None})
                            this_rows_cells.append({"merge":None})
                            # inject the row of cells into the table model
                            child_model['table']['rows'].append({"cells":this_rows_cells})
                        # Add the child table to the parent table
                        parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_load_balancers}}", table)

def add_load_balancer_targets_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # First populate non-vpc Load Balancer Targers
    for region, attributes in topology.items():
        if not region in non_region_topology_keys and "non_vpc_lb_target_groups" in attributes.keys() and attributes['non_vpc_lb_target_groups']:
            parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"Heading 2","text":f"Region {region} (Non VPC Targets)"}]}]})
            for rownum, tg in enumerate(sorted(attributes['non_vpc_lb_target_groups'], key=lambda d : d['TargetType']), start=1):
                if rownum > 1: # Inject space between Target Group tables
                    parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":""}]}]})
                # Build the child table
                child_model = deepcopy(word_table_models.lb_target_group_tbl)
                child_model['table']['rows'][0]['cells'][0]['paragraphs'].append({"style":"regularbold","text":f"TARGET GROUP ARN: {tg['TargetGroupArn']}"})
                child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":tg['TargetGroupName']})
                child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":"<NA>"})
                child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":tg['TargetType']})
                child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":tg['LoadBalancerArns']})
                child_model['table']['rows'][4]['cells'][0]['paragraphs'].append({"style":"No Spacing","text":"<NA>"})
                child_model['table']['rows'][4]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":"<NA>"})
                child_model['table']['rows'][4]['cells'][2]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthyThresholdCount'])})
                child_model['table']['rows'][4]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":str(tg['UnhealthyThresholdCount'])})
                child_model['table']['rows'][4]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthCheckTimeoutSeconds'])})
                child_model['table']['rows'][4]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthCheckIntervalSeconds'])})
                for rownum2, target in enumerate(tg['HealthChecks'], start=1):
                    this_rows_cells = []
                    # Shade every other row for readability
                    if not (rownum2 % 2) == 0:
                        row_color = alternating_row_color
                    else:
                        row_color = None
                    # Build word table rows & cells
                    health_reason = "---" if target['TargetHealth']['State'] == "healthy" else target['TargetHealth']['Reason']
                    health_description = "---" if target['TargetHealth']['State'] == "healthy" else target['TargetHealth']['Description']
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":target['Target']['Id']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":"<NA>"}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":target['TargetHealth']['State']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":health_reason}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":health_description}]})
                    this_rows_cells.append({"merge":None})
                    # inject the row of cells into the table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Add the child table to the parent table
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name == ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name}"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                if not vpc['lb_target_groups']:
                    parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":"No Target Groups Present"}]}]})
                else:
                    for rownum, tg in enumerate(sorted(vpc['lb_target_groups'], key=lambda d : d['TargetType']), start=1):
                        if rownum > 1: # Inject space between Target Group tables
                            parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":""}]}]})
                        # Build the child table
                        child_model = deepcopy(word_table_models.lb_target_group_tbl)
                        child_model['table']['rows'][0]['cells'][0]['paragraphs'].append({"style":"regularbold","text":f"TARGET GROUP ARN: {tg['TargetGroupArn']}"})
                        child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":tg['TargetGroupName']})
                        child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":f"{tg['Protocol']}:{tg['Port']}"})
                        child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":tg['TargetType']})
                        child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":tg['LoadBalancerArns']})
                        child_model['table']['rows'][4]['cells'][0]['paragraphs'].append({"style":"No Spacing","text":tg['HealthCheckProtocol']})
                        child_model['table']['rows'][4]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":str(tg['Port'])})
                        child_model['table']['rows'][4]['cells'][2]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthyThresholdCount'])})
                        child_model['table']['rows'][4]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":str(tg['UnhealthyThresholdCount'])})
                        child_model['table']['rows'][4]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthCheckTimeoutSeconds'])})
                        child_model['table']['rows'][4]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":str(tg['HealthCheckIntervalSeconds'])})
                        for rownum2, target in enumerate(tg['HealthChecks'], start=1):
                            this_rows_cells = []
                            # Shade every other row for readability
                            if not (rownum2 % 2) == 0:
                                row_color = alternating_row_color
                            else:
                                row_color = None
                            # Build word table rows & cells
                            health_reason = "---" if target['TargetHealth']['State'] == "healthy" else target['TargetHealth']['Reason']
                            health_description = "---" if target['TargetHealth']['State'] == "healthy" else target['TargetHealth']['Description']
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":target['Target']['Id']}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(target['Target']['Port'])}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":target['TargetHealth']['State']}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":health_reason}]})
                            this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":health_description}]})
                            this_rows_cells.append({"merge":None})
                            # inject the row of cells into the table model
                            child_model['table']['rows'].append({"cells":this_rows_cells})
                        # Add the child table to the parent table
                        parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_lb_target_groups}}", table)

def add_instances_to_word_doc():
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in filtered_topology.items():
        if not vpcs:
            pass
        else:
            for vpc in vpcs:
                this_parent_tbl_rows_cells = []
                try:
                    vpc_name = [tag['Value'] for tag in vpc['Tags'] if tag['Key'] == "Name"][0]
                except KeyError:
                    # Object has no name
                    vpc_name = ""
                except IndexError:
                    vpc_name = ""
                # Create the parent table row and cells
                this_parent_tbl_rows_cells.append({"paragraphs":[{"style":"Heading 2","text":f"Region: {region} / VPC: {vpc_name} ({len(vpc['ec2_instances'])} Instances)"}]})
                # inject the row of cells into the table model
                parent_model['table']['rows'].append({"cells":this_parent_tbl_rows_cells})
                # Build the child table
                if not vpc['ec2_instances']:
                    parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":"No EC2 Instances"}]}]})
                else:
                    for rownum, inst in enumerate(vpc['ec2_instances'], start=1):
                        if rownum > 1: # inject space between instance tables
                            parent_model['table']['rows'].append({"cells":[{"paragraphs":[{"style":"No Spacing","text":""}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                        child_model = deepcopy(word_table_models.ec2_inst_tbl)
                        try: # Get Instance name
                            inst_name = [tag['Value'] for tag in inst['Tags'] if tag['Key'] == "Name"][0]
                        except KeyError:
                            # Object has no name
                            inst_name = ""
                        except IndexError:
                            # Object has no name
                            inst_name = ""
                        try: # Get Public IP Address
                            public_ip = inst['PublicIpAddress']
                        except KeyError:
                            public_ip = ""
                        # Build word table rows & cells
                        child_model['table']['rows'][0]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst_name})
                        child_model['table']['rows'][0]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['ImageId']})
                        child_model['table']['rows'][0]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":inst['InstanceType']})
                        child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst['Placement']['AvailabilityZone']})
                        child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['PrivateIpAddress']})
                        child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":public_ip})
                        child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst['PlatformDetails']})
                        child_model['table']['rows'][2]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['Architecture']})
                        child_model['table']['rows'][2]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":inst['State']['Name']})
                        # Add network interfaces to table
                        inst_label = inst_name if not inst_name == "" else inst['InstanceId']
                        child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs": [{"style": "regularbold", "text": f"{inst_label} NETWORK INTERFACES"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                        if not inst['NetworkInterfaces']:
                            child_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Network Interfaces"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
                        else:
                            child_model['table']['rows'].append(word_table_models.ec2_inst_interface_tbl_header)
                            for rownum2, intf in enumerate(sorted(inst['NetworkInterfaces'], key=lambda d : d['Attachment']['DeviceIndex']), start=1):
                                this_rows_cells = []
                                # Shade every other row for readability
                                if not (rownum2 % 2) == 0:
                                    row_color = alternating_row_color
                                else:
                                    row_color = None
                                try: # Get Public IP if applicable
                                    public_ip = intf['Association']['PublicIp']
                                except KeyError:
                                    public_ip = ""
                                # Get Security Groups
                                sec_grps = [sg['GroupId'] for sg in intf['Groups'] if sg['GroupId'].startswith("sg-")]
                                # Add interface cells to table row
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":f"{intf['NetworkInterfaceId']}/{intf['Description']}"}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":[ipadd['PrivateIpAddress'] for ipadd in sorted(intf['PrivateIpAddresses'], key = lambda d : d['Primary'])]}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":public_ip}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":f"{intf['SubnetId']}({get_subnet_name_by_id(intf['SubnetId'],vpc)})"}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":sec_grps}]})
                                this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":str(intf['Attachment']['DeviceIndex'])}]})
                                # inject cells into the child table row
                                child_model['table']['rows'].append({"cells":this_rows_cells})
                        # Add the child table to the parent table
                        parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPCs Present"}]}]})
    else:
        table = build_table(doc_obj, parent_model)
        replace_placeholder_with_table(doc_obj, "{{py_ec2_inst}}", table)

if __name__ == "__main__":
    try:
        if not args.skip_topology:
            ec2 = boto3.client('ec2', verify=False)
            available_regions = get_regions()
            topology = {}
            try: 
                account_alias = boto3.client('iam', verify=False).list_account_aliases()['AccountAliases'][0]
            except IndexError:
                account_alias = ""
            topology['account'] = {
                "id": boto3.client('sts', verify=False).get_caller_identity().get('Account'),
                "alias": account_alias
            }

            add_regions_to_topology()

            rprint("\n[yellow]STEP 1/12: DISCOVER REGION VPCS")
            add_vpcs_to_topology()

            rprint("\n\n[yellow]STEP 2/12: DISCOVER VPC NETWORK ELEMENTS")
            add_network_elements_to_vpcs()

            rprint("\n[yellow]STEP 3/12: DISCOVER REGION PREFIX LISTS")
            add_prefix_lists_to_topology()

            rprint("\n[yellow]STEP 4/12: DISCOVER REGION VPN CUSTOMER GATEWAYS")
            add_vpn_customer_gateways_to_topology()

            rprint("\n[yellow]STEP 5/12: DISCOVER REGION VPN CONNECTIONS ATTACHED TO TRANSIT GATEAWAYS")
            add_vpn_tgw_connections_to_topology()

            rprint("\n[yellow]STEP 6/12: DISCOVER REGION VPC ENDPOINT SERVICES")
            add_endpoint_services_to_topology()

            rprint("\n\n[yellow]STEP 7/12: DISCOVERING ACCOUNT VPC PEERING CONNECTIONS")
            add_vpc_peering_connections_to_topology()

            rprint("\n\n[yellow]STEP 8/12: DISCOVERING REGION TRANSIT GATEWAYS")
            add_transit_gateways_to_topology()

            rprint("\n\n[yellow]STEP 9/12: DISCOVERING REGION TRANSIT GATEWAY ROUTES")
            add_transit_gateway_routes_to_topology()

            rprint("\n\n[yellow]STEP 10/12: DISCOVERING DIRECT CONNECT")
            add_direct_connect_to_topology()
        else:
            # Get the first toplogy file from the current working directory
            fp = pathlib.Path(os.getcwd())
            file_list = [f.name for f in fp.iterdir() if f.is_file() and f.name.endswith(".json")]
            if len(file_list) > 1:
                rprint("\n\n :x: [red]Multiple Topology files detected in the current working directory.")
                rprint("[red]Please ensure only one exists. Exiting...")
                sys.exit(1)
            else:
                with open(file_list[0], "r") as f:
                    topology = json.load(f)

        filtered_topology = {region:attributes['vpcs'] for region, attributes in topology.items() if not region in non_region_topology_keys}

        rprint("\n\n[yellow]STEP 11/12: BUILD WORD DOCUMENT OBJECT")
        doc_obj = create_word_obj_from_template(word_template)
        rprint("[yellow]    Creating VPC table...")
        add_vpcs_to_word_doc()
        rprint("[yellow]    Creating Subnets table...")
        add_subnets_to_word_doc()
        rprint("[yellow]    Creating Route Tables table...")
        add_route_tables_to_word_doc()
        rprint("[yellow]    Creating Route Table Routes table...")
        add_routes_to_word_doc()
        rprint("[yellow]    Creating Prefix Lists table...")
        add_prefix_lists_to_word_doc()
        rprint("[yellow]    Creating Network ACLs table...")
        add_network_acls_to_word_doc()
        rprint("[yellow]    Creating Network ACL Inbound Entries table...")
        add_netacl_inbound_entries_to_word_doc()
        rprint("[yellow]    Creating Network ACL Outbound Entries table...")
        add_netacl_outbound_entries_to_word_doc()
        rprint("[yellow]    Creating Security Groups table...")
        add_security_groups_to_word_doc()
        rprint("[yellow]    Creating Security Group Inbound Entries table...")
        add_sg_inbound_entries_to_word_doc()
        rprint("[yellow]    Creating Security Group Outbound Entries table...")
        add_sg_outbound_entries_to_word_doc()
        rprint("[yellow]    Creating Internet Gateways table...")
        add_internet_gateways_to_word_doc()
        rprint("[yellow]    Creating Egress-Only Internet Gateways table...")
        add_egress_only_internet_gateways_to_word_doc()
        rprint("[yellow]    Creating NAT Gateways table...")
        add_nat_gateways_to_word_doc()
        rprint("[yellow]    Creating Endpoint Services table...")
        add_endpoint_services_to_word_doc()
        rprint("[yellow]    Creating Endpoints table...")
        add_endpoints_to_word_doc()
        rprint("[yellow]    Creating VPC Peerings table...")
        add_vpc_peerings_to_word_doc()
        rprint("[yellow]    Creating Transit Gateways table...")
        add_transit_gateways_to_word_doc()
        rprint("[yellow]    Creating Transit Gateway Routes table...")
        add_transit_gateway_routes_to_word_doc()
        rprint("[yellow]    Creating VPN Customer Gateways table...")
        add_vpn_customer_gateways_to_word()
        rprint("[yellow]    Creating VPN Transit Gateway Connections table...")
        add_vpn_tgw_connections_to_word()
        rprint("[yellow]    Creating VPN Gateways table...")
        add_vpn_gateways_to_word_doc()
        rprint("[yellow]    Creating Direct Connect Gateways table...")
        add_direct_connect_gateways_to_word_doc()
        rprint("[yellow]    Creating Load Balancers table...")
        add_load_balancers_to_word_doc()
        rprint("[yellow]    Creating Load Balancer Target Groups table...")
        add_load_balancer_targets_to_word_doc()
        rprint("[yellow]    Creating EC2 Instances table...")
        add_instances_to_word_doc()

        rprint("\n\n[yellow]STEP 12/12: WRITING ARTIFACTS TO FILE SYSTEM")
        rprint("    [yellow]Saving Word document...")
        # Get Platform
        system_os = platform.system().lower()
        def slasher():
            # Returns the correct file system slash for the detected platform
            return "\\" if system_os == "windows" else "/"
        if topology['account']['alias'] == "":
            word_file = f"{os.getcwd()}{slasher()}{topology_folder}{slasher()}{topology['account']['id']} {str(datetime.datetime.now()).split()[0].replace('-','')}.docx"
        else:
            word_file = f"{os.getcwd()}{slasher()}{topology_folder}{slasher()}{topology['account']['alias']} {str(datetime.datetime.now()).split()[0].replace('-','')}.docx"
        try:
            doc_obj.save(word_file)
        except:
            rprint(f"\n\n:x: [red]Could not save output to {word_file}. If it is open please close and try again.\n\n")
            sys.exit()
        if not args.skip_topology:
            if topology['account']['alias'] == "":
                topology_file = f"{os.getcwd()}{slasher()}{topology_folder}{slasher()}{topology['account']['id']} {str(datetime.datetime.now()).split()[0].replace('-','')}.json"
            else:
                topology_file = f"{os.getcwd()}{slasher()}{topology_folder}{slasher()}{topology['account']['alias']} {str(datetime.datetime.now()).split()[0].replace('-','')}.json"
            rprint("    [yellow]Saving raw AWS topology...")
            with open(topology_file, "w") as f:
                f.write(json.dumps(topology,indent=4,default=datetime_converter))

        rprint(f"\n\n[green]FILES WRITTEN, ALL DONE!!!!")
        rprint(f"    [green]AWS As-Built Word Document written to: [blue]{word_file}")
        if not args.skip_topology:
            rprint(f"    [green]Raw AWS topology file written to: [blue]{topology_file}")
        rprint("[yellow]NOTE: Be sure to update the Word Document Table of Contents as dynamically-created headlines will not be reflected in the TOC until that is done.\n\n")
    except KeyboardInterrupt:
        rprint("\n\n[red]Exiting due to keyboard interrupt...\n")
        sys.exit()