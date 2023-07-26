import xmltodict
import os
import pandas as pd

def take_info(name_archive, values):
    with open(f'nfs/{name_archive}', 'rb') as archive_xml:
        dic_archive = xmltodict.parse(archive_xml)

        if 'NFe' in dic_archive:
            infos_nf = dic_archive["NFe"]["infNFe"]
        else:
                infos_nf = dic_archive['nfeProc']["NFe"]["infNFe"]
        nf_number = infos_nf['@Id']
        emp_emissor = infos_nf['emit']['xNome']
        client_name = infos_nf['dest']['xNome']
        address = infos_nf['dest']['enderDest']
        if 'vol' in infos_nf['transp']:
            weight = infos_nf['transp']['vol']['pesoB']
        else:
            weight = 'Peso não informado'
        values.append([nf_number, emp_emissor, client_name, address, weight])



list_archives = os.listdir("nfs")

columns = ["nf_number", "emp_emissor", "client_name", "address", "weight"]
values = []

for archive in list_archives:
    take_info(archive, values)

table = pd.DataFrame(columns=columns, data=values)
table.to_excel('NotasFiscais.xlsx', index=False)

print('Sucesso!')