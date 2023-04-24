import boto3, botocore.exceptions, requests, sys, datetime, json, os, argparse, pathlib, datetime, platform
from rich import print as rprint
from docx import Document
from dcnet_msofficetools.docx_extensions import build_table, replace_placeholder_with_table
from copy import deepcopy
import word_table_models
from urllib3.exceptions import InsecureRequestWarning
requests.packages.urllib3.disable_warnings(category=InsecureRequestWarning)

# SET COMMANDLINE ARGUMENT PARAMETERS
parser = argparse.ArgumentParser()
parser.add_argument(
    '-t',
    action='store_true',
    default=False,
    dest='skip_topology',
    help='Skip bulding topology file from AWS API (use existing JSON topology files in working directory)'
    )
args = parser.parse_args()

# GLOBAL VARIABLES
output_verbosity = 0   # 0 (Default) or 1 (Verbose)
topology_folder = "topologies"
word_template = "template.docx"
table_header_color = "506279" # Dark Blue
green_spacer = "8FD400" # CC Green/Lime
red_spacer = "F12938" # CC Red
orange_spacer = "FF7900" # CC Orange
alternating_row_color = "D5DCE4" # Light Blue
region_list = [] # Leave blank to auto-pull and check all regions
aws_protocol_map = { # Maps AWS protocol numbers to user-friendly names
    "-1": "All Traffic",
    "6": "TCP",
    "17": "UDP",
    "1": "ICMPv4",
    "58": "ICMPv6"
}
# non_region_topology_keys = ["account", "vpc_peering_connections", "direct_connect_gateways"]

# HELPER FUNCTIONS
def datetime_converter(obj):
    # Converts datetime objects to a string timestamp. Needed to render JSON
    if isinstance(obj, datetime.datetime):
        return obj.__str__()

def create_word_obj_from_template(template_file):
    # Attempt to read a Word document from filesystem and load it as a docx object
    try:
        return Document(template_file)
    except:
        rprint(f"\n\n:x: [red]Could not open [blue]{template_file}[red]. Please make sure it exists and is a valid Microsoft Word document. Exiting...")
        sys.exit(1)

def extract_name_from_aws_tags(obj):
    # Extract the name from a given object by searching the object's list of AWS tags
    try:
        name = [tag['Value'] for tag in obj['Tags'] if tag['Key'] == "Name"][0]
    except KeyError:
        # Object has no name
        name = "<unnamed>"
    except IndexError:
        # Object has no name
        name = "<unnamed>"
    return name

def get_subnet_name_by_id(source_subnet_id, source_vpc=None):
    # Return a subnet's name tag (if available) by looking up the name in the topology VPC structure
    if source_vpc:
        for subnet in source_vpc['subnets']:
            if subnet['SubnetId'] == source_subnet_id:
                subnet_name = extract_name_from_aws_tags(subnet)
                break
    else:
        subnet_name = None
        for vpcs in region_vpcs.values():
            for vpc in vpcs:
                for subnet in vpc['subnets']:
                    if subnet['SubnetId'] == source_subnet_id:
                        subnet_name = extract_name_from_aws_tags(subnet)
                        break
                if subnet_name:
                    break
    subnet_name = "" if not subnet_name else subnet_name
    return subnet_name
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

        topology['regions'][region] = {
            "vpcs": [],
            "prefix_lists": [],
            "customer_gateways": [],
            "vpn_tgw_connections": [],
            "endpoint_services": [],
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
    for region in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for VPCs...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        elb = boto3.client('elbv2',region_name=region,verify=False)
        try:
            response = ec2.describe_vpcs()['Vpcs']
            topology['regions'][region]["non_vpc_lb_target_groups"] = [tg for tg in elb.describe_target_groups()['TargetGroups'] if not "VpcId" in tg.keys()]
            for tg in topology['regions'][region]["non_vpc_lb_target_groups"]:
                tg['HealthChecks'] = elb.describe_target_health(TargetGroupArn=tg['TargetGroupArn'])['TargetHealthDescriptions']
            for vpc in response:
                is_empty_default_vpc = fingerprint_vpc(region, vpc, ec2)
                if not is_empty_default_vpc:
                    topology['regions'][region]['vpcs'].append(vpc)
        except botocore.exceptions.ClientError:
            rprint(f":x: [red]Client Error reported for region {region}. Most likely no VPCs exist, continuing...")

def add_network_elements_to_vpcs():
    for k, v in topology['regions'].items():
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
    for region in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for Prefix Lists...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        try:
            pls = [pl for pl in ec2.describe_prefix_lists()['PrefixLists']]
            topology['regions'][region]['prefix_lists'] = pls
        except botocore.exceptions.ClientError:
            rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_vpn_customer_gateways_to_topology():
    for region in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for Customer Gateways...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        try:
            cgws = [cgw for cgw in ec2.describe_customer_gateways()['CustomerGateways']]
            topology['regions'][region]['customer_gateways'] = cgws
        except botocore.exceptions.ClientError:
            rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_vpn_tgw_connections_to_topology():
    for region in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for VPN Connections Attached to Transit Gateways...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        try:
            tgw_vpns = [conn for conn in ec2.describe_vpn_connections()['VpnConnections'] if "TransitGatewayId" in conn.keys()]
            topology['regions'][region]['vpn_tgw_connections'] = tgw_vpns
        except botocore.exceptions.ClientError:
            rprint(f":x: [red]Client Error reported for region {region}. Skipping...")

def add_endpoint_services_to_topology():
    for region, v in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for Endpoint Services...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        try:
            ep_svcs = [svc for svc in ec2.describe_vpc_endpoint_services()['ServiceDetails'] if not svc['Owner'] == "amazon"]
            topology['regions'][region]['endpoint_services'] = ep_svcs
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
    for region in topology['regions']:
        rprint(f"    [yellow]Interrogating Region {region} for Transit Gateways...")
        ec2 = boto3.client('ec2',region_name=region,verify=False)
        try:
            tgws = [tgw for tgw in ec2.describe_transit_gateways()['TransitGateways']]
            for tgw in tgws:
                attachments = [attachment for attachment in ec2.describe_transit_gateway_attachments()['TransitGatewayAttachments'] if attachment['TransitGatewayId'] == tgw['TransitGatewayId']]
                for attachment in attachments: # Loop through VPC attachments and set ApplianceModeSupport option
                    appliance_mode_support = "disable"
                    if attachment['ResourceType'] == "vpc":
                        attachment_options = ec2.describe_transit_gateway_vpc_attachments(Filters=[{'Name':'transit-gateway-attachment-id','Values':[attachment['TransitGatewayAttachmentId']]}])['TransitGatewayVpcAttachments'][0]
                        appliance_mode_support = attachment_options['Options']['ApplianceModeSupport']
                        attachment['SubnetIds'] = attachment_options['SubnetIds']
                    else:
                        attachment['SubnetIds'] = ["<NA>"]
                    attachment['ApplianceModeSupport'] = appliance_mode_support
                tgw['attachments'] = attachments
                rts = [rt for rt in ec2.describe_transit_gateway_route_tables()['TransitGatewayRouteTables'] if rt['TransitGatewayId'] == tgw['TransitGatewayId']]
                tgw['route_tables'] = rts
            topology['regions'][region]['transit_gateways'] = tgws
        except botocore.exceptions.ClientError as e:
            if "(UnauthorizedOperation)" in str(e):
                rprint(f"[red]Unauthorized Operation reported while pulling Transit Gateways from {region}. Skipping...")
            else:
                print(e)

def add_transit_gateway_routes_to_topology():
    for region in topology['regions']:
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
                    "TransitGatewayRouteTableName": extract_name_from_aws_tags(rt),
                    "Routes":routes
                })
        except botocore.exceptions.ClientError as e:
            if "(UnauthorizedOperation)" in str(e):
                rprint(f"[red]Unauthorized Operation reported while pulling Transit Gateway Route Tables from {region}. Skipping...")
            else:
                print(e)
        topology['regions'][region]['transit_gateway_routes'] = tgw_routes

# BEST PRACTICE CHECK FUNCTIONS
def add_transit_gateway_best_practice_analysis_to_word_doc(doc_obj):
    def run_tgw_quantity_check(tgws):
        test_description = "Because Transit Gateways are highly available by design, multiple gateways per region are not required for high availability. https://docs.aws.amazon.com/vpc/latest/tgw/tgw-best-design-practices.html"
        regions_with_multiple_tgws = [f"Region {region} has {len(gateways)} Transit Gateways." for tgw in tgws for region, gateways in tgw.items() if len(gateways) > 1]
        if regions_with_multiple_tgws:
            test_status = "warning"
            regions_with_multiple_tgws.append("Confirm any specific use cases exist that call for multiple Transit Gateways in a region.")
            test_results = regions_with_multiple_tgws
        else:
            test_status = "pass"
            test_results = "All regions have no more than one Transit Gateway."
        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    def run_unique_bgp_asn_check(tgws):
        test_description = "For deployments with multiple transit gateways, a unique Autonomous System Number (ASN) for each transit gateway is recommended. https://docs.aws.amazon.com/vpc/latest/tgw/tgw-best-design-practices.html"
        if len(tgws) == 1:
            test_status = "not-applicable"
            test_results = "Single Transit Gateway detected. This test is not applicable."
        else:
            tgw_bgp_asns = [gateway['Options']['AmazonSideAsn'] for tgw in tgws for gateways in tgw.values() for gateway in gateways]
            if len(tgw_bgp_asns) == len(set(tgw_bgp_asns)):
                test_status = "pass"
                test_results = "All Transit Gateway ASNs are unique."
            else:
                # Get the duplicates
                duplicates = set()
                for i, asn1 in enumerate(tgw_bgp_asns):
                    for asn2 in tgw_bgp_asns[i+1:]:
                        if asn1 == asn2:
                            duplicates.add(asn1)
                test_status = "fail"
                test_results = f"Detected re-use of ASNs: {list(duplicates)}"
        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    def run_one_net_acl_check(tgws):
        test_description = "Create one network ACL and associate it with all of the subnets that are associated with the transit gateway. https://docs.aws.amazon.com/vpc/latest/tgw/tgw-best-design-practices.html"
        # Build a list of Transit Gateway ID / Subnet, key / value pairs
        subnets = [{gateway['TransitGatewayId']:subnet} for tgw in tgws for gateways in tgw.values() for gateway in gateways for attachment in gateway['attachments'] if "SubnetIds" in attachment.keys() for subnet in attachment['SubnetIds']]
        # Build a list of Subnet IDs for each Transit Gateway
        tgw_subnets = {}
        for each in subnets:
            for tgw, subnet in each.items():
                if tgw in tgw_subnets:
                    tgw_subnets[tgw].append(subnet)
                else:
                    tgw_subnets[tgw] = [subnet]
        # Loop through the Transit Gateways and pull the Network ACL assigned to each subnet. Update fail_list if any VPCs have attachments in different NACLS
        fail_list = []
        account_net_acls_in_use = {}
        for tgw, subnets in tgw_subnets.items():
            network_acls_in_use = {}
            for subnet in subnets:
                # Loop through all VPC Network ACLs and pull the Network ACL ID where this subnet is associated
                found = False
                for vpcs in region_vpcs.values():
                    for vpc in vpcs:
                        for netacl in vpc['network_acls']:
                            for association in netacl['Associations']:
                                if association['SubnetId'] == subnet:
                                    if not vpc['VpcId'] in network_acls_in_use.keys(): # VPC Not yet in dictionary
                                        network_acls_in_use[vpc['VpcId']] = [netacl['NetworkAclId']]
                                    elif not netacl['NetworkAclId'] in network_acls_in_use[vpc['VpcId']]: # VPC in dictionary but NetACL not in list yet
                                        network_acls_in_use[vpc['VpcId']].append(netacl['NetworkAclId'])
                                    found = True
                                    break
                        if found:
                            break
                    if found:
                            break
            account_net_acls_in_use[tgw] = network_acls_in_use
            # Check to see if any VPCs have multiple Net ACLs, if so add it to the failed list
            for vpc, nacl_list in network_acls_in_use.items():
                if len(nacl_list) > 1: # Best Practice Check Failure
                    fail_list.append({
                        "TransitGatewayId": tgw,
                        "vpcs": [
                            {
                                "VpcId": vpc,
                                "NetworkAcls": nacl_list
                            }
                        ]
                    })
        
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All Transit Gateway VPC Attachment Subnets are in a single Network ACL."
        else:
            test_status = "fail"
            test_results = [
                "The following Transit Gateways have Attachment Subnets in multiple Network ACLs:",   
            ] + [tgw['TransitGatewayId'] + '/' + vpc['VpcId'] + '/NetACLs: ' + str(vpc['NetworkAcls']) for tgw in fail_list for vpc in tgw['vpcs']]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results,
            "net_acls": account_net_acls_in_use
        }

    def run_net_acl_open_check(test_results):
        test_description = "Create one network ACL and associate it with all of the subnets that are associated with the transit gateway. Keep the network ACL open in both the inbound and outbound directions. https://docs.aws.amazon.com/vpc/latest/tgw/tgw-best-design-practices.html"
        fail_list = []
        for tgw, vpc_dict in test_results['net_acls'].items():
            for vpc_id, nacl_list in vpc_dict.items():
                vpc_dict = [{"region":region,"vpc":this_vpc} for region, vpcs in region_vpcs.items() for this_vpc in vpcs if this_vpc['VpcId'] == vpc_id][0]
                for nacl in nacl_list:
                    ingress_entries = [entry for acl in vpc_dict['vpc']['network_acls'] if acl['NetworkAclId'] == nacl for entry in sorted(acl['Entries'], key = lambda d : d['RuleNumber']) if not entry['Egress']]
                    egress_entries = [entry for acl in vpc_dict['vpc']['network_acls'] if acl['NetworkAclId'] == nacl for entry in sorted(acl['Entries'], key = lambda d : d['RuleNumber']) if entry['Egress']]
                    ingress_cidr_block = ingress_entries[0]['CidrBlock']
                    ingress_protocol = ingress_entries[0]['Protocol']
                    ingress_action =  ingress_entries[0]['RuleAction']
                    if not ingress_cidr_block == "0.0.0.0/0" or not ingress_protocol == "-1" or not ingress_action == "allow": # Test Failed, add to fail_list
                        fail_list.append({
                            "TransitGatewayId": tgw,
                            "Region": vpc_dict['region'],
                            "VpcId": vpc_id,
                            "NetworkAclId": nacl,
                            "Direction": "Ingress",
                            "EntryRuleNumber": ingress_entries[0]['RuleNumber']
                        })
                    egress_cidr_block = egress_entries[0]['CidrBlock']
                    egress_protocol = egress_entries[0]['Protocol']
                    egress_action =  egress_entries[0]['RuleAction']
                    if not egress_cidr_block == "0.0.0.0/0" or not egress_protocol == "-1" or not egress_action == "allow": # Test Failed, add to fail_list
                        fail_list.append({
                            "TransitGatewayId": tgw,
                            "Region": vpc_dict['region'],
                            "VpcId": vpc_id,
                            "NetworkAclId": nacl,
                            "Direction": "Egress",
                            "EntryRuleNumber": ingress_entries[0]['RuleNumber']
                        })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All Transit Gateway VPC Attachment Subnet Network ACLs are open."
        else:
            test_status = "fail"
            test_results = [
                "The following Transit Gateways VPC Attachment Subnet ACLs don't appear to be open. Please review their configuration to confirm:",   
            ] + [nacl['VpcId']+ '/' + nacl['NetworkAclId'] + '(' + nacl['Direction'] + ')' for nacl in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    def run_vpn_attachment_bgp_check(tgws):
        test_description = "Use Border Gateway Protocol (BGP) Site-to-Site VPN connections. https://docs.aws.amazon.com/vpc/latest/tgw/tgw-best-design-practices.html"
        fail_list = []
        vpn_attachments = [{"TransitGatewayId":gw['TransitGatewayId'],"VpnId":attachment['ResourceId']} for each in tgws for gws in each.values() for gw in gws for attachment in gw['attachments'] if attachment['ResourceType'] == "vpn"]
        for each in vpn_attachments:
            each['vpn_connections'] = [conn for region, attributes in topology.items() if not region in non_region_topology_keys and "vpn_tgw_connections" in attributes for conn in attributes['vpn_tgw_connections'] if conn['VpnConnectionId'] == each['VpnId']]
        for attachment in vpn_attachments:
            for vpn in attachment['vpn_connections']:
                # Check if attachments have static routes (infer BGP disabled if so)
                if vpn['Options']['StaticRoutesOnly']:
                    fail_list.append({
                        "TransitGatewayId": vpn['TransitGatewayId'],
                        "VpnId": vpn['VpnConnectionId'],
                        "Routing": "Static",
                        "Tunnels": vpn['VgwTelemetry']
                    })
                elif sum([tunnel['AcceptedRouteCount'] for tunnel in vpn['VgwTelemetry']]) == 0: # No learned routes, infer BGP not enabled or not functioning
                    fail_list.append({
                        "TransitGatewayId": vpn['TransitGatewayId'],
                        "VpnId": vpn['VpnConnectionId'],
                        "Routing": "Dynamic",
                        "Tunnels": vpn['VgwTelemetry']
                    })

        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All Transit Gateway VPN connections are learning routes dynamically (BGP enabled)."
        else:
            test_status = "fail"
            test_results = [
                "The following Transit Gateway VPNs are not enabled for dynamic routing or are not learning routes:",   
            ] + [vpn['TransitGatewayId'] + '/' + vpn['VpnId'] for vpn in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    tgws = [{region:attributes['transit_gateways']} for region, attributes in topology['regions'].items() if attributes['transit_gateways']]
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)

    tgw_results = {
        "passed": 0,
        "failed": 0
    }
    if tgws:
        # Run best practice checks
        tgw_quantity_check = run_tgw_quantity_check(tgws)
        if tgw_quantity_check['status'] == "pass":
            tgw_results['passed'] += 1
        else:
            tgw_results['failed'] += 1
        unique_bgp_asn_check = run_unique_bgp_asn_check(tgws)
        if unique_bgp_asn_check['status'] in ["pass","not-applicable"]:
            tgw_results['passed'] += 1
        else:
            tgw_results['failed'] += 1
        one_net_acl_check = run_one_net_acl_check(tgws)
        if one_net_acl_check['status'] == "pass":
            tgw_results['passed'] += 1
        else:
            tgw_results['failed'] += 1
        net_acl_open_check = run_net_acl_open_check(one_net_acl_check)
        if net_acl_open_check['status'] == "pass":
            tgw_results['passed'] += 1
        else:
            tgw_results['failed'] += 1
        vpn_attachment_bgp_check = run_vpn_attachment_bgp_check(tgws)
        if vpn_attachment_bgp_check['status'] == "pass":
            tgw_results['passed'] += 1
        else:
            tgw_results['failed'] += 1

        # Create the tgw_quantity_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Populate the child model with test data
        header_color = green_spacer if tgw_quantity_check['status'] == "pass" else orange_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"TGW PER-REGION QUANTITY CHECK: {tgw_quantity_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = tgw_quantity_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = tgw_quantity_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

        # Create the unique_bgp_asn_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Inject a space between child tables
        parent_model['table']['rows'].append({"cells":[]})
        # Populate the child model with test data
        header_color = green_spacer if unique_bgp_asn_check['status'] in ["pass","not-applicable"] else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"UNIQUE BGP ASN CHECK: {unique_bgp_asn_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = unique_bgp_asn_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = unique_bgp_asn_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

        # Create the one_net_acl_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Inject a space between child tables
        parent_model['table']['rows'].append({"cells":[]})
        # Populate the child model with test data
        header_color = green_spacer if one_net_acl_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"ONE NET ACL FOR ATTACHMENT SUBNETS CHECK: {one_net_acl_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = one_net_acl_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = one_net_acl_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

        # Create the net_acl_open_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Inject a space between child tables
        parent_model['table']['rows'].append({"cells":[]})
        # Populate the child model with test data
        header_color = green_spacer if net_acl_open_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"NET ACL OPEN CHECK: {net_acl_open_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = net_acl_open_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = net_acl_open_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

        # Create the vpn_attachment_bgp_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Inject a space between child tables
        parent_model['table']['rows'].append({"cells":[]})
        # Populate the child model with test data
        header_color = green_spacer if vpn_attachment_bgp_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"VPN ATTACHMENT BGP CHECK: {vpn_attachment_bgp_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = vpn_attachment_bgp_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = vpn_attachment_bgp_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

    # Write parent table to Word
    if not parent_model['table']['rows']: # Completely Empty Table (no Transit Gateways)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Transit Gateways Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_tgw_best_practices}}", table)
    return tgw_results

def add_vpn_best_practice_analysis_to_word_doc(doc_obj):
    def run_vpn_tunnel_status_check():
        test_description = "Report any VPN tunnel connections in the DOWN state."
        fail_list = []
        for vpn in vpns:
            for tunnel in vpn['VgwTelemetry']:
                if tunnel['Status'] == "DOWN":
                    fail_list.append({
                        "VpnId": vpn['VpnConnectionId'],
                        "TunnelIp": tunnel['OutsideIpAddress']
                    })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All VPN Tunnel Connections are in Status 'UP' State."
        else:
            test_status = "fail"
            test_results = [
                "The following VPN Tunnel Connections are in Status 'DOWN' State:",   
            ] + [vpn['VpnId'] + '/' + vpn['TunnelIp'] for vpn in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    # Get VPN connections in vpn_tgw_connections dictionary key
    vpns = [vpn for region, attributes in topology.items() if region not in non_region_topology_keys and "vpn_tgw_connections" in attributes for vpn in attributes['vpn_tgw_connections']]
    # Add VPN connections in VPC VPN Gateways
    vpns += [vpn for vpcs in region_vpcs.values() for vpc in vpcs for vgw in vpc['vpn_gateways'] for vpn in vgw['connections']]
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)

    vpn_results = {
        "passed": 0,
        "failed": 0
    }
    if vpns:
        # Run best practice checks
        vpn_tunnel_status_check = run_vpn_tunnel_status_check()
        if vpn_tunnel_status_check['status'] == "pass":
            vpn_results['passed'] += 1
        else:
            vpn_results['failed'] += 1

        # Create the _vpn_tunnel_status_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Populate the child model with test data
        header_color = green_spacer if vpn_tunnel_status_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"VPN TUNNEL STATUS CHECK: {vpn_tunnel_status_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = vpn_tunnel_status_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = vpn_tunnel_status_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

    # Write parent table to Word
    if not parent_model['table']['rows']: # Completely Empty Table (no VPNs)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPNs Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpn_health}}", table)
    return vpn_results

def add_vpc_best_practice_analysis_to_word_doc(doc_obj):
    def run_empty_vpc_check():
        test_description = "Report any VPCs with no EC2 instances."
        fail_list = []
        for region, vpcs in region_vpcs.items():
            for vpc in vpcs:
                if not vpc['ec2_instances']:
                    fail_list.append({
                        "region": region,
                        "vpc_id": vpc['VpcId'],
                        "vpc_name": extract_name_from_aws_tags(vpc)
                    })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All VPCs have at least one EC2 instance."
        else:
            test_status = "warn"
            test_results = [
                "The following VPCs have no EC2 instances. Review VPC use case to confirm this is intentional:",   
            ] + [vpc['region'] + '/' + vpc['vpc_id'] + '(' + vpc['vpc_name'] + ')' for vpc in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }
   
    def run_multi_az_check():
        test_description = "When you add subnets to your VPC to host your application, create them in multiple Availability Zones. https://docs.aws.amazon.com/vpc/latest/userguide/vpc-security-best-practices.html"
        fail_list = []
        for region, vpcs in region_vpcs.items():
            for vpc in vpcs:
                availability_zones = list(set([subnet['AvailabilityZone'] for subnet in vpc['subnets']]))
                if len(availability_zones) == 1:
                    fail_list.append({
                        "region": region,
                        "vpc_id": vpc['VpcId'],
                        "vpc_name": extract_name_from_aws_tags(vpc)
                    })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All VPCs have subnets in two or more Availability Zones."
        else:
            test_status = "fail"
            test_results = [
                "The following VPCs have subnets in only one Availability Zone:",   
            ] + [vpc['region'] + '/' + vpc['vpc_id'] + '(' + vpc['vpc_name'] + ')' for vpc in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)

    vpc_results = {
        "passed": 0,
        "failed": 0
    }
    if len([vpc['VpcId'] for region, vpcs in region_vpcs.items() for vpc in vpcs]) > 0:
        # Run best practice checks
        empty_vpc_check = run_empty_vpc_check()
        if empty_vpc_check['status'] == "pass":
            vpc_results['passed'] += 1
        else:
            vpc_results['failed'] += 1
        multi_az_check = run_multi_az_check()
        if multi_az_check['status'] == "pass":
            vpc_results['passed'] += 1
        else:
            vpc_results['failed'] += 1

        # Create the empty_vpc_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Populate the child model with test data
        header_color = green_spacer if empty_vpc_check['status'] == "pass" else orange_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"EMPTY VPC CHECK: {empty_vpc_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = empty_vpc_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = empty_vpc_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

        # Create the multi_az_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Inject a space between child tables
        parent_model['table']['rows'].append({"cells":[]})
        # Populate the child model with test data
        header_color = green_spacer if multi_az_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"MULTI-AZ VPC CHECK: {multi_az_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = multi_az_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = multi_az_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

    # Write parent table to Word
    if not parent_model['table']['rows']: # Completely Empty Table (no VPNs)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No VPNs Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpc_health}}", table)
    return vpc_results

def add_lb_best_practice_analysis_to_word_doc(doc_obj):
    def run_lb_target_health_check():
        test_description = "Report any Load Balancers with Targets in Unhealthy State."
        fail_list = []
        for lbtg in lbtgs:
            for target in lbtg['lbtg']['HealthChecks']:
                if not target['TargetHealth']['State'] == "healthy":
                    # Get VPC Name
                    for vpcs in region_vpcs.values():
                        for vpc in vpcs:
                            if vpc['VpcId'] == lbtg['lbtg']['VpcId']:
                                vpc_name = extract_name_from_aws_tags(vpc)
                    fail_list.append({
                        "region": lbtg['region'],
                        "vpc_id": lbtg['lbtg']['VpcId'],
                        "vpc_name": vpc_name,
                        "target_id": target['Target']['Id']
                    })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All load balancer targets are in a healthy state."
        else:
            test_status = "fail"
            test_results = [
                "The following Load Balancer Targets are in an unhealthy state:",   
            ] + [target['region'] + '/' + target['vpc_id'] + '(' + target['vpc_name'] + ')/' + target['target_id'] for target in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    lbtgs = [{"region":region,"lbtg":lbtg} for region, vpcs in region_vpcs.items() for vpc in vpcs for lbtg in vpc['lb_target_groups']]

    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)

    lb_results = {
        "passed": 0,
        "failed": 0
    }
    if len([vpc['VpcId'] for region, vpcs in region_vpcs.items() for vpc in vpcs]) > 0:
        # Run best practice checks
        lb_target_health_check = run_lb_target_health_check()
        if lb_target_health_check['status'] == "pass":
            lb_results['passed'] += 1
        else:
            lb_results['failed'] += 1

        # Create the lb_target_health_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Populate the child model with test data
        header_color = green_spacer if lb_target_health_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"LB TARGET HEALTH CHECK: {lb_target_health_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = lb_target_health_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = lb_target_health_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

    # Write parent table to Word
    if not parent_model['table']['rows']: # Completely Empty Table (no VPNs)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Load Balancer Targets Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_lb_health}}", table)
    return lb_results

def add_ec2_best_practice_analysis_to_word_doc(doc_obj):
    def run_ec2_ena_enabled_check():
        test_description = "Amazon EC2 provides enhanced networking capabilities through the Elastic Network Adapter (ENA). To use enhanced networking, you must install the required ENA module and enable ENA support. https://docs.aws.amazon.com/AWSEC2/latest/UserGuide/enhanced-networking-ena.html"
        fail_list = []
        for instance in instances:
            if not "EnaSupport" in instance['instance'] or not instance['instance']['EnaSupport']:
                fail_list.append({
                    "region": instance['region'],
                    "vpc_id": instance['vpc_id'],
                    "vpc_name": instance['vpc_name'],
                    "instance_id": instance['instance']['InstanceId'],
                    "instance_name": extract_name_from_aws_tags(instance['instance'])
                })
        # Create Best Practice Check Return Status
        if not fail_list:
            test_status = "pass"
            test_results = "All EC2 instances have ENA Support enabled."
        else:
            test_status = "fail"
            test_results = [
                "The following EC2 Instances do not have ENA Support enabled:",   
            ] + [instance['region'] + '/' + instance['vpc_id'] + '/' + instance['instance_id'] + '(' + instance['instance_name'] + ')' for instance in fail_list]

        return {
            "description": test_description,
            "status": test_status,
            "results": test_results
        }

    instances = [{"region":region,"vpc_id":vpc['VpcId'],"vpc_name":extract_name_from_aws_tags(vpc),"instance":instance} for region, vpcs in region_vpcs.items() for vpc in vpcs for instance in vpc['ec2_instances']]
    
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)

    ec2_results = {
        "passed": 0,
        "failed": 0
    }
    if instances:
        # Run best practice checks
        ec2_ena_enabled_check = run_ec2_ena_enabled_check()
        if ec2_ena_enabled_check['status'] == "pass":
            ec2_results['passed'] += 1
        else:
            ec2_results['failed'] += 1

        # Create the ec2_ena_enabled_check child table model
        child_model = deepcopy(word_table_models.best_practices_tbl)
        # Populate the child model with test data
        header_color = green_spacer if ec2_ena_enabled_check['status'] == "pass" else red_spacer
        child_model['table']['rows'][0]['cells'][0]['background'] = header_color
        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"ENA SUPPORT ENABLED CHECK: {ec2_ena_enabled_check['status'].upper()}"
        child_model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = ec2_ena_enabled_check['description']
        child_model['table']['rows'][2]['cells'][1]['paragraphs'][0]['text'] = ec2_ena_enabled_check['results']
        # Inject child model into parent model
        parent_model['table']['rows'].append({"cells":[child_model]})

    # Write parent table to Word
    if not parent_model['table']['rows']: # Completely Empty Table (no VPNs)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No EC2 Instances Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_inst_health}}", table)
    return ec2_results
# BUILD WORD TABLE FUNCTIONS
def add_vpcs_to_word_doc(doc_obj):
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
        vpc_name = extract_name_from_aws_tags(vpc['vpc'])
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

def add_route_tables_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_rts}}", table)

def add_routes_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_rt_routes}}", table)

def add_prefix_lists_to_word_doc(doc_obj):
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_prefix_lists}}", table)

def add_subnets_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_subnets}}", table)

def add_network_acls_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_netacls}}", table)

def add_netacl_inbound_entries_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_netacl_in_entries}}", table)

def add_netacl_outbound_entries_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_netacl_out_entries}}", table)

def add_security_groups_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_sgs}}", table)

def add_sg_inbound_entries_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_sg_in_entries}}", table)

def add_sg_outbound_entries_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_sg_out_entries}}", table)

def add_internet_gateways_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_igws}}", table)

def add_egress_only_internet_gateways_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_eigws}}", table)

def add_nat_gateways_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_ngws}}", table)

def add_endpoint_services_to_word_doc(doc_obj):
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
                ep_svc_name = extract_name_from_aws_tags(epsvc)
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
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Endpoint Services Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_endpoint_services}}", table)

def add_endpoints_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_endpoints}}", table)

def add_vpc_peerings_to_word_doc(doc_obj):
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

def add_transit_gateways_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, attributes in topology.items():
        if not region in non_region_topology_keys and attributes['transit_gateways']: # Ignore these dictionary keys, they are not a region, also don't run if there are no transit gateways in the region
            # Create Table title (Region)
            parent_model['table']['rows'].append({"cells": [{"paragraphs":[{"style":"Heading 2","text":f"Region: {region}"}]}]})
            for rownum, tgw in enumerate(attributes['transit_gateways']):
                if rownum > 0: # Inject an empty row to space the data
                    parent_model['table']['rows'].append({"cells":[]})
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
                child_model['table']['rows'][1]['cells'][4]['paragraphs'].append({"style":"No Spacing","text":tgw['TransitGatewayId']})
                child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":str(tgw['Options']['AmazonSideAsn'])})
                child_model['table']['rows'][2]['cells'][4]['paragraphs'].append({"style":"No Spacing" if tgw['OwnerId'] == topology['account']['id'] else "redtext","text":tgw['OwnerId']})
                # Populate child table model with spacer and attachment header
                child_model['table']['rows'].append({"cells":[{"background":green_spacer,"paragraphs":[{"style":"regularbold","text":"ATTACHMENTS"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
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
                    # Build Subnet Ids including names
                    if len(attch['SubnetIds']) == 1 and attch['SubnetIds'][0] == "<NA>":
                        subnets = attch['SubnetIds']
                    else:
                        subnets = [subnet + "(" + get_subnet_name_by_id(subnet) + ")" for subnet in attch['SubnetIds']]
                    # Add data to row/cells
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch_name}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch['TransitGatewayAttachmentId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":resource_type}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":attch['ResourceId']}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":rt_id}]})
                    this_rows_cells.append({"background":row_color,"paragraphs":[{"style":"No Spacing","text":subnets}]})
                    # add attachment data row to child table model
                    child_model['table']['rows'].append({"cells":this_rows_cells})
                # Populate child table model with spacer and route table header
                child_model['table']['rows'].append({"cells":[{"background":red_spacer,"paragraphs":[{"style":"regularbold","text":"ROUTE TABLES"}]},{"merge":None},{"merge":None},{"merge":None},{"merge":None},{"merge":None}]})
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
                parent_model['table']['rows'].append({"cells":[child_model]})
    # Model has been build, now convert it to a python-docx Word table object
    if not parent_model['table']['rows']: # Completely Empty Table (no VPCs at all)
        parent_model['table']['rows'].append({"cells":[{"paragraphs": [{"style": "No Spacing", "text": "No Transit Gateways Present"}]}]})
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_tgws}}", table)

def add_transit_gateway_routes_to_word_doc(doc_obj):
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_tgw_routes}}", table)

def add_vpn_customer_gateways_to_word(doc_obj):
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpn_cgws}}", table)

def add_vpn_tgw_connections_to_word(doc_obj):
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpn_s2s}}", table)

def add_vpn_gateways_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_vpn_vpgs}}", table)

def add_direct_connect_gateways_to_word_doc(doc_obj):
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

def add_load_balancers_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
                                tg_names = [tg['TargetGroupArn'].split('/')[1] + '' if not "Weight" in tg.keys() else tg['TargetGroupArn'].split('/')[1] + '(Weight: ' + str(tg['Weight']) + ')' for tg in listener['DefaultActions'][0]['ForwardConfig']['TargetGroups']]
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_load_balancers}}", table)

def add_load_balancer_targets_to_word_doc(doc_obj):
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
    for region, vpcs in region_vpcs.items():
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_lb_target_groups}}", table)

def add_instances_to_word_doc(doc_obj):
    # Create the parent table model
    parent_model = deepcopy(word_table_models.parent_tbl)
    # Populate the table model with data
    for region, vpcs in region_vpcs.items():
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
                        try: # GET ENA SUPPORT
                            ena_support = "YES" if inst['EnaSupport'] else "NO"
                        except KeyError:
                            ena_support = "Not Compatible"
                        # Build word table rows & cells
                        child_model['table']['rows'][0]['cells'][0]['paragraphs'][0]['text'] = f"INSTANCE ID: {inst['InstanceId']}"
                        child_model['table']['rows'][1]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst_name})
                        child_model['table']['rows'][1]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['ImageId']})
                        child_model['table']['rows'][1]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":f"{inst['InstanceType']}/{ena_support}"})
                        child_model['table']['rows'][2]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst['Placement']['AvailabilityZone']})
                        child_model['table']['rows'][2]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['PrivateIpAddress']})
                        child_model['table']['rows'][2]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":public_ip})
                        child_model['table']['rows'][3]['cells'][1]['paragraphs'].append({"style":"No Spacing","text":inst['PlatformDetails']})
                        child_model['table']['rows'][3]['cells'][3]['paragraphs'].append({"style":"No Spacing","text":inst['Architecture']})
                        child_model['table']['rows'][3]['cells'][5]['paragraphs'].append({"style":"No Spacing","text":inst['State']['Name']})
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
    table = build_table(doc_obj, parent_model)
    replace_placeholder_with_table(doc_obj, "{{py_ec2_inst}}", table)

def build_word_document():
    rprint("\n\n[yellow]STEP 11/14: BUILD WORD DOCUMENT OBJECT")
    doc_obj = create_word_obj_from_template(word_template)
    rprint("[yellow]    Creating VPC table...")
    add_vpcs_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Subnets table...")
    add_subnets_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Route Tables table...")
    add_route_tables_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Route Table Routes table...")
    add_routes_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Prefix Lists table...")
    add_prefix_lists_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Network ACLs table...")
    add_network_acls_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Network ACL Inbound Entries table...")
    add_netacl_inbound_entries_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Network ACL Outbound Entries table...")
    add_netacl_outbound_entries_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Security Groups table...")
    add_security_groups_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Security Group Inbound Entries table...")
    add_sg_inbound_entries_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Security Group Outbound Entries table...")
    add_sg_outbound_entries_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Internet Gateways table...")
    add_internet_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Egress-Only Internet Gateways table...")
    add_egress_only_internet_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating NAT Gateways table...")
    add_nat_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Endpoint Services table...")
    add_endpoint_services_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Endpoints table...")
    add_endpoints_to_word_doc(doc_obj)
    rprint("[yellow]    Creating VPC Peerings table...")
    add_vpc_peerings_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Transit Gateways table...")
    add_transit_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Transit Gateway Routes table...")
    add_transit_gateway_routes_to_word_doc(doc_obj)
    rprint("[yellow]    Creating VPN Customer Gateways table...")
    add_vpn_customer_gateways_to_word(doc_obj)
    rprint("[yellow]    Creating VPN Transit Gateway Connections table...")
    add_vpn_tgw_connections_to_word(doc_obj)
    rprint("[yellow]    Creating VPN Gateways table...")
    add_vpn_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Direct Connect Gateways table...")
    add_direct_connect_gateways_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Load Balancers table...")
    add_load_balancers_to_word_doc(doc_obj)
    rprint("[yellow]    Creating Load Balancer Target Groups table...")
    add_load_balancer_targets_to_word_doc(doc_obj)
    rprint("[yellow]    Creating EC2 Instances table...")
    add_instances_to_word_doc(doc_obj)
    return doc_obj

def perform_best_practices_analysis(doc_obj):
    rprint("\n\n[yellow]STEP 12/14: PERFORM BEST PRACTICE ANALYSIS")
    rprint("[yellow]    Performing Transit Gateway Best Practices/Health Analysis and writing to Word table...")
    tgw_results = add_transit_gateway_best_practice_analysis_to_word_doc(doc_obj)
    rprint("[yellow]    Performing VPN Best Practices/Health Analysis and writing to Word table...")
    vpn_results = add_vpn_best_practice_analysis_to_word_doc(doc_obj)
    rprint("[yellow]    Performing VPC Best Practices/Health Analysis and writing to Word table...")
    vpc_results = add_vpc_best_practice_analysis_to_word_doc(doc_obj)
    rprint("[yellow]    Performing Load Balancer Best Practices/Health Analysis and writing to Word table...")
    lb_results = add_lb_best_practice_analysis_to_word_doc(doc_obj)
    rprint("[yellow]    Performing EC2 Instance Best Practices/Health Analysis and writing to Word table...")
    ec2_results = add_ec2_best_practice_analysis_to_word_doc(doc_obj)
    return {
        "tgw": tgw_results,
        "vpn": vpn_results,
        "vpc": vpc_results,
        "lb": lb_results,
        "ec2": ec2_results
    }

def create_account_dashboard(doc_obj, analysis_results):
    rprint("\n\n[yellow]STEP 13/14: CREATE ACCOUNT DASHBOARD")
    model = deepcopy(word_table_models.account_dashboard_tbl)
    regions_in_use = [region for region, attributes in topology.items() if not region in non_region_topology_keys and (attributes['vpcs'] or attributes['transit_gateways'])]     
    vpc_count = len([vpc['VpcId'] for vpcs in region_vpcs.values() for vpc in vpcs])
    ec2_count = len([inst['InstanceId'] for vpcs in region_vpcs.values() for vpc in vpcs for inst in vpc['ec2_instances']])
    model['table']['rows'][0]['cells'][1]['paragraphs'][0]['text'] = topology['account']['id']
    model['table']['rows'][0]['cells'][4]['paragraphs'][0]['text'] = topology['account']['alias']
    model['table']['rows'][1]['cells'][1]['paragraphs'][0]['text'] = regions_in_use
    model['table']['rows'][1]['cells'][4]['paragraphs'][0]['text'] = f"{ec2_count} EC2 instances across {vpc_count} VPCs"
    analysis_cells = []
    tgw_background = green_spacer if analysis_results['tgw']['failed'] == 0 else red_spacer
    analysis_cells.append({"background":tgw_background, "paragraphs":[{"style":"regularbold","text": f"{str(analysis_results['tgw']['passed'])} of {str(analysis_results['tgw']['passed'] + analysis_results['tgw']['failed'])} checks passed."}]})
    analysis_cells.append({"merge":None})
    vpn_background = green_spacer if analysis_results['vpn']['failed'] == 0 else red_spacer
    analysis_cells.append({"background":vpn_background, "paragraphs":[{"style":"regularbold","text": f"{str(analysis_results['vpn']['passed'])} of {str(analysis_results['vpn']['passed'] + analysis_results['vpn']['failed'])} checks passed."}]})
    vpc_background = green_spacer if analysis_results['vpc']['failed'] == 0 else red_spacer
    analysis_cells.append({"background":vpc_background, "paragraphs":[{"style":"regularbold","text": f"{str(analysis_results['vpc']['passed'])} of {str(analysis_results['vpc']['passed'] + analysis_results['vpc']['failed'])} checks passed."}]})
    lb_background = green_spacer if analysis_results['lb']['failed'] == 0 else red_spacer
    analysis_cells.append({"background":lb_background, "paragraphs":[{"style":"regularbold","text": f"{str(analysis_results['lb']['passed'])} of {str(analysis_results['lb']['passed'] + analysis_results['lb']['failed'])} checks passed."}]})
    ec2_background = green_spacer if analysis_results['ec2']['failed'] == 0 else red_spacer
    analysis_cells.append({"background":ec2_background, "paragraphs":[{"style":"regularbold","text": f"{str(analysis_results['ec2']['passed'])} of {str(analysis_results['ec2']['passed'] + analysis_results['ec2']['failed'])} checks passed."}]})
    model['table']['rows'].append({"cells":analysis_cells})
    table = build_table(doc_obj, model)
    replace_placeholder_with_table(doc_obj, "{{py_account_dashboard}}", table)

def write_artifacts_to_filesystem(doc_obj):
    rprint(f"\n\n[yellow]STEP 14/14: WRITING ARTIFACTS TO FILE SYSTEM FOR {topology['account']['alias']}")
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

if __name__ == "__main__":
    try:
        if not args.skip_topology: # If we don't supply the -t flag, we need to build the topology by actively reaching out to the AWS API
            ec2 = boto3.client('ec2', verify=False)
            available_regions = get_regions() # Pull all regions the account has access to
            topology = { # Create the Account topology skeleton
                "account": None,
                "regions": {},
                "vpc_peering_connections": [],
                "direct_connect_gateways": []
            }
            try: # Pull the account alias
                account_alias = boto3.client('iam', verify=False).list_account_aliases()['AccountAliases'][0]
            except IndexError: # No alias configured
                account_alias = ""
            topology['account'] = {
                "id": boto3.client('sts', verify=False).get_caller_identity().get('Account'),
                "alias": account_alias
            }

            add_regions_to_topology()

            rprint("\n[yellow]STEP 1/14: DISCOVER REGION VPCS")
            add_vpcs_to_topology()

            rprint("\n\n[yellow]STEP 2/14: DISCOVER VPC NETWORK ELEMENTS")
            add_network_elements_to_vpcs()

            rprint("\n[yellow]STEP 3/14: DISCOVER REGION PREFIX LISTS")
            add_prefix_lists_to_topology()

            rprint("\n[yellow]STEP 4/14: DISCOVER REGION VPN CUSTOMER GATEWAYS")
            add_vpn_customer_gateways_to_topology()

            rprint("\n[yellow]STEP 5/14: DISCOVER REGION VPN CONNECTIONS ATTACHED TO TRANSIT GATEAWAYS")
            add_vpn_tgw_connections_to_topology()

            rprint("\n[yellow]STEP 6/14: DISCOVER REGION VPC ENDPOINT SERVICES")
            add_endpoint_services_to_topology()

            rprint("\n\n[yellow]STEP 7/14: DISCOVERING ACCOUNT VPC PEERING CONNECTIONS")
            add_vpc_peering_connections_to_topology()

            rprint("\n\n[yellow]STEP 8/14: DISCOVERING REGION TRANSIT GATEWAYS")
            add_transit_gateways_to_topology()

            rprint("\n\n[yellow]STEP 9/14: DISCOVERING REGION TRANSIT GATEWAY ROUTES")
            add_transit_gateway_routes_to_topology()

            rprint("\n\n[yellow]STEP 10/14: DISCOVERING DIRECT CONNECT")
            add_direct_connect_to_topology()
            topologies = [topology]
        else: # -t flag supplied so we know the topology already exists in a JSON file or files
            # Build a list of topology files
            fp = pathlib.Path(os.getcwd())
            file_list = [f.name for f in fp.iterdir() if f.is_file() and f.name.endswith(".json")]
            # Extract all topologies from file system and store in a list
            topologies = []
            for file in file_list:
                with open(file, "r") as f:
                    topologies.append(json.load(f))

        # The topology is ready, now start parsing the topology data model to render Word tables
        for topology in topologies:
            # Build a dictionary of just regions and their VPCs (so we don't have to do a nested loop every time we want to loop over region VPCs)
            region_vpcs = {region:attributes['vpcs'] for region, attributes in topology['regions'].items()}

            # Populate the Word document with all the configuration data
            doc_obj = build_word_document()

            analysis_results = perform_best_practices_analysis(doc_obj)

            create_account_dashboard(doc_obj, analysis_results)

            write_artifacts_to_filesystem(doc_obj)
    except KeyboardInterrupt:
        rprint("\n\n[red]Exiting due to keyboard interrupt...\n")
        sys.exit()