# import json, xlwt
import xlsxwriter

d = {
        "type": "Microsoft.Network/networkSecurityGroups/securityRules",
        "apiVersion": "2019-06-01",
        "name": "[concat(parameters('networkSecurityGroups_dc01_nsg_name'), '/SecurityCenter-JITRule_-1282941078_D94F68C391F84FFFB715F439EA7EBF4E')]",
        "properties": {
            "provisioningState": "Succeeded",
            "description": "ASC JIT Network Access rule for policy 'default' of VM 'dc01'.",
            "protocol": "*",
            "sourcePortRange": "*",
            "destinationPortRange": "5985",
            "sourceAddressPrefix": "*",
            "destinationAddressPrefix": "10.0.1.4",
            "access": "Deny",
            "priority": 4095,
            "direction": "Inbound",
            "sourcePortRanges": "abc",
            "destinationPortRanges": "def",
            "sourceAddressPrefixes": "ghi",
            "destinationAddressPrefixes": "yxz"
            }
}

workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()
row = 0
col = 0

for key, value in d.items():
    if type(value) is dict:
        col += 1
        for key1, value1 in value.items():
            col += 1
            worksheet.write(row, col, key1)
            worksheet.write(row + 1, col, value1)
            print(value1)
    else:
        col += 1
        worksheet.write(row, col, key)
        worksheet.write(row + 1, col, value)

workbook.close()












# print(d['properties']['description'])
# workbook = xlsxwriter.Workbook('data.xlsx')
# worksheet = workbook.add_worksheet()
# row = 0
# col = 0
# col2 = 0
#
# for key in d.keys():
#     col += 1
#     worksheet.write(row, col, key)
#     print(key)
# for item in d.values():
#     col2 += 1
#     worksheet.write(row + 1, col2, item)
#     print(item)
#
# workbook.close()