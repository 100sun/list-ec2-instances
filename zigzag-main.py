properties = [
    "InstanceId",
    "Tag)Name",
    "InstanceState",
    "LaunchTime",
    "InstanceType",
    "PlatformType",
    "VolumeId",
    "VolumeSize(Gib)",
    "VpcId",
    "VpcName",
    "SubnetId",
    "KeyPair",
    "PrivateIpAddress",
    "PublicIpAddress",
    "Tag)karpenter.sh/provisioner-name",
    "Tag)service",
    "Tag)Application",
    "Tag)Phase",
    "Tag)Terraform",
    "Tag)aws:autoscaling:groupName",
    "Tag)aws:cloudformation:logical-id",
    "Tag)aws:cloudformation:stack-id",
    "Tag)aws:cloudformation:stack-name",
    "Tag)aws:ec2spot:fleet-request-id",
    # "Tag)eks:cluster-name",
    # "Tag)eks:nodegroup-name",
    # "Tag)kubernetes.io/cluster/zigzag-production-batch",
    # "Tag)kubernetes.io/cluster/zigzag-production-1",
    # "Tag)kubernetes.io/cluster/zigzag-production-2",
    # "Tag)kubernetes.io/cluster/zigzag-production-3",
    # "Tag)kubernetes.io/cluster/zigzag-production-5",
    # "Tag)kubernetes.io/cluster/zigzag-production-6",
    # "Tag)kubernetes.io/cluster/zigzag-production-10",
    # "Tag)kubernetes.io/cluster/beta-cluster-20220209",
    # "Tag)kubernetes.io/cluster/production-cluster-20220516",
    # "Tag)k8s.io/cluster-autoscaler/enabled",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-batch",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-1",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-2",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-3",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-5",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-6",
    # "Tag)k8s.io/cluster-autoscaler/zigzag-production-10",
    # "Tag)k8s.io/cluster-autoscaler/beta-cluster-20220209",
    # "Tag)k8s.io/cluster-autoscaler/production-cluster-20220516",
    "Tag)etc..",
]

from pprint import pprint
from collections import Counter

all_tags = dict()


def get_ec2_instances(account_name):
    import boto3
    import json
    from datetime import datetime

    boto3.setup_default_session(
        profile_name=f"{account_name}", region_name="ap-northeast-2"
    )
    ec2_client = boto3.client("ec2")

    def get_refined_instance_tags(tags):
        refined_tags = {}
        for tag in tags:
            refined_tags[tag["Key"]] = tag["Value"]
        return refined_tags

    def get_vpc_name(vpc_id):
        if not vpc_id:
            return
        for vpc_tag in ec2_client.describe_vpcs(VpcIds=[vpc_id])["Vpcs"][0]["Tags"]:
            if vpc_tag["Key"] == "Name":
                return vpc_tag["Value"]

    def get_ebs_size(ebs_id):
        if not ebs_id:
            return
        return (
            ec2_client.describe_volumes(VolumeIds=[ebs_id])
                .get("Volumes", [])[0]
                .get("Size", "")
        )

    instances_info = []
    for instance in ec2_client.describe_instances()["Reservations"]:
        instance = instance["Instances"][0]
        instance_tags = get_refined_instance_tags(instance["Tags"])
        if instance["State"]["Name"] != "running" or instance_tags.get("Name", "").startswith(
                "karpenter") or instance_tags.get("aws:autoscaling:groupName", "").startswith("eks"):
            continue
        ebs_id = (
            instance.get("BlockDeviceMappings", [])[0]
                .get("Ebs", "")
                .get("VolumeId", "")
        )
        vpc_id = instance.get("VpcId", "")

        for k, v in instance_tags.items():
            if k not in all_tags.keys():
                all_tags[k] = []
            else:
                all_tags[k].append(v)
        # all_tags.update(Counter(instance_tags))
        # all_tags.update(instance_tags)
        # all_tags.update(Counter(instance_tags))
        basic_info = [
            instance["InstanceId"],
            instance_tags.pop("Name", ""),
            instance["State"]["Name"],
            instance["LaunchTime"].strftime("%m/%d/%Y, %H:%M:%S"),
            instance["InstanceType"],
            instance.get("PlatformDetails", ""),
            ebs_id,
            get_ebs_size(ebs_id),
            vpc_id,
            get_vpc_name(vpc_id),
            instance.get("SubnetId", ""),
            instance.get("KeyName", ""),
            instance.get("PrivateIpAddress", ""),
            instance.get("PublicIpAddress", ""),
            instance_tags.pop("Service", ""),
        ]
        for i in range(properties.index("Tag)service"), len(properties) - 1):
            basic_info.append(instance_tags.pop(properties[i][4:], ""))
        basic_info.append(json.dumps(instance_tags))
        instances_info.append(basic_info)

    return instances_info


def write_to_csv(account_name):
    import csv

    with open(f"{account_name}.csv", "w", encoding="UTF8", newline="") as f:
        writer = csv.writer(f)

        writer.writerow(properties)
        writer.writerows(get_ec2_instances(account_name))


def write_to_xlsx(account_name):
    import xlsxwriter

    workbook = xlsxwriter.Workbook(f"{account_name}.xlsx", {"remove_timezone": True})
    worksheet = workbook.add_worksheet(f"{account_name}")

    instances_info = get_ec2_instances(account_name)
    for i in range(len(properties)):
        worksheet.write(0, i, properties[i])

    for i in range(0, len(instances_info)):
        for j in range(len(instances_info[i])):
            if j == 0:
                worksheet.write_url(
                    i + 1,
                    j,
                    f"https://ap-northeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-northeast-2#InstanceDetails:instanceId={instances_info[i][j]}",
                    string=instances_info[i][j],
                )
            elif j == 6:
                worksheet.write_url(
                    i + 1,
                    j,
                    f"https://ap-northeast-2.console.aws.amazon.com/ec2/v2/home?region=ap-northeast-2#VolumeDetails:volumeId={instances_info[i][j]}",
                    string=instances_info[i][j],
                )
            elif j == 8:
                worksheet.write_url(
                    i + 1,
                    j,
                    f"https://ap-northeast-2.console.aws.amazon.com/vpc/home?region=ap-northeast-2#VpcDetails:VpcId={instances_info[i][j]}",
                    string=instances_info[i][j],
                )
            else:
                worksheet.write(i + 1, j, instances_info[i][j])

    workbook.close()


if __name__ == "__main__":
    write_to_xlsx("zigzag-main")
    # sorted(all_tags.keys(), key=lambda x: len(all_tags[x]), reverse=True)
