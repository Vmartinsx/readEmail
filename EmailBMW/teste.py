import re


teste =  'chassi: 52636276372'

m = re.search('chassi: (.*)$', teste, re.MULTILINE)
print(m)
print(m.group(1))
print(type(m))


nome_criacao_pasta = str(m.group(1)).rstrip()


