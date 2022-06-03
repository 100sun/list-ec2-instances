properties = [
    "InstanceId",
    "Tag)Name",
    "InstanceType",
    "InstanceState",
    "LaunchTime",
    "KeyPair",
    "VpcName",
    "SubnetId",
    "VpcId",
    "PrivateIpAddress",
    "PublicIpAddress",
    "Tag)Service",
    "Tag)autoscaling:groupName",
    "Tag)ec2:fleet-id",
    "Tag)cloudformation:logical-id",
    "Tag)cloudformation:stack-id",
    "Tag)cloudformation:stack-name",
    "Tag)etc.."
]


def get_ec2_instances(account_name):
    import boto3
    import json

    boto3.setup_default_session(profile_name=f"{account_name}", region_name='ap-northeast-2')
    ec2_client = boto3.client('ec2')

    def get_refined_instance_tags(tags):
        refined_tags = {}
        for tag in tags:
            refined_tags[tag["Key"]] = tag["Value"]
        return refined_tags

    def get_vpc_name(vpc_id):
        if not vpc_id:
            return
        for vpc_tag in ec2_client.describe_vpcs(VpcIds=[vpc_id])['Vpcs'][0]['Tags']:
            if vpc_tag['Key'] == 'Name':
                return vpc_tag['Value']

    instances_info = []
    for instance in ec2_client.describe_instances()["Reservations"]:
        instance = instance["Instances"][0]
        instance_tags = get_refined_instance_tags(instance['Tags'])
        instances_info.append([
            instance["InstanceId"],
            instance_tags.pop("Name", ""),
            instance["InstanceType"],
            instance["State"]["Name"],
            instance["LaunchTime"],
            instance.get("KeyName", ""),
            get_vpc_name(instance.get("VpcId", "")),
            instance.get("SubnetId", ""),
            instance.get("VpcId", ""),
            instance.get("PrivateIpAddress", ""),
            instance.get("PublicIpAddress", ""),
            instance_tags.pop("Service", ""),
            instance_tags.pop("aws:autoscaling:groupName", ""),
            instance_tags.pop("aws:ec2:fleet-id", ""),
            instance_tags.pop("aws:cloudformation:logical-id", ""),
            instance_tags.pop("aws:cloudformation:stack-id", ""),
            instance_tags.pop("aws:cloudformation:stack-name", ""),
            json.dumps(instance_tags),
        ])

    return instances_info


def write_to_csv(account_name):
    import csv
    with open(f'{account_name}.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)

        writer.writerow(properties)
        writer.writerows(get_ec2_instances(account_name))


def write_to_xlsx(account_name):
    import xlsxwriter

    workbook = xlsxwriter.Workbook(f"{account_name}.xlsx", {'remove_timezone': True})
    worksheet = workbook.add_worksheet(f"{account_name}")

    instances_info = get_ec2_instances(account_name)
    for i in range(len(properties)):
        worksheet.write(0, i, properties[i])

    for i in range(0, len(instances_info)):
        for j in range(len(instances_info[i])):
            if j == 0:
                worksheet.write_url(i + 1, j,
                                    f'https://ap-northeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-northeast-2#InstanceDetails:instanceId={instances_info[i][j]}',
                                    string=instances_info[i][j])
            # elif j == 4:
            #     worksheet.write_datetime(i + 1, j, instances_info[i][j])
            elif j == 8:
                worksheet.write_url(i + 1, j,
                                    f'https://ap-northeast-2.console.aws.amazon.com/vpc/home?region=ap-northeast-2#VpcDetails:VpcId={instances_info[i][j]}',
                                    string=instances_info[i][j])
            else:
                worksheet.write(i + 1, j, instances_info[i][j])

    workbook.close()


if __name__ == "__main__":
    accounts = ["croquis-main", "zigzag-main", "zigzag-alpha"]
    write_to_xlsx(accounts[0])
    # write_to_xlsx(accounts[1])
    # write_to_xlsx(accounts[2])
