def get_ec2_instances(account_name):
    import boto3

    boto3.setup_default_session(profile_name=f"{account_name}", region_name='ap-northeast-2')
    ec2_client = boto3.client('ec2')

    def get_refined_instance_tags(tags):
        refined_tags = {}
        for tag in tags:
            refined_tags[tag["Key"]] = tag["Value"]
        return refined_tags

    def get_vpc_name(vpc_id):
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
            get_vpc_name(instance["VpcId"]),
            instance["SubnetId"],
            instance["VpcId"],
            instance["PrivateIpAddress"],
            instance.get("PublicIpAddress"),
            instance_tags.pop("Service", ""),
            instance_tags.pop("aws:autoscaling:groupName", ""),
            instance_tags.pop("aws:ec2:fleet-id", ""),
            instance_tags.pop("aws:cloudformation:logical-id", ""),
            instance_tags.pop("aws:cloudformation:stack-id", ""),
            instance_tags.pop("aws:cloudformation:stack-name", ""),
            instance_tags,
        ])

    return instances_info


def write_to_csv(account_name):
    import csv
    with open(f'ec2_instances_{account_name}.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)

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
        writer.writerow(properties)
        writer.writerows(get_ec2_instances(account_name))


if __name__ == "__main__":
    accounts = ["croquis-main", "zigzag-main", "zigzag-alpha"]
    for account in accounts[:1]:
        write_to_csv(account)
