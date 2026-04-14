import requests
import pandas as pd
import time

base = "https://www3.tjrj.jus.br/consultasportal/api/v1/telefonesEnderecos/serventias1Inst/listar?recaptcha=skipRecap2020&codigoOrgao={comarca}&codigoRegional=0&codigoTipoServentia=1&codigoAtribuicao=999"
lista_final = []
comarcas = {
    "CAPITAL": "406",
    "ANGRA DOS REIS": "383",
    "ARARUAMA": "341",
    "ARRAIAL DO CABO": "1022",
    "BARRA DO PIRAÍ": "385",
    "BARRA MANSA": "384",
    "BELFORD ROXO": "1870",
    "BOM JARDIM": "343",
    "BOM JESUS DE ITABAPOANA": "342",
    "BÚZIOS": "2246",
    "CABO FRIO": "386",
    "CACHOEIRAS DE MACACU": "344",
    "CAMBUCI-SÃO JOSÉ DE UBÁ": "345",
    "CAMPOS DOS GOYTACAZES": "387",
    "CANTAGALO": "346",
    "CARAPEBUS / QUISSAMÃ": "2371",
    "CARMO": "347",
    "CASIMIRO DE ABREU": "348",
    "CONCEIÇÃO DE MACABU": "349",
    "CORDEIRO-MACUCO": "350",
    "DUAS BARRAS": "351",
    "DUQUE DE CAXIAS": "388",
    "ENGENHEIRO PAULO DE FRONTIN": "352",
    "GUAPIMIRIM": "2219",
    "IGUABA GRANDE": "2189",
    "ITABORAÍ": "389",
    "ITAGUAÍ": "390",
    "ITALVA-CARDOSO MOREIRA": "2249",
    "ITAOCARA": "353",
    "ITAPERUNA": "391",
    "ITATIAIA": "2339",
    "JAPERI": "2342",
    "LAJE DO MURIAÉ": "354",
    "MACAÉ": "392",
    "MAGÉ": "393",
    "MANGARATIBA": "355",
    "MARICÁ": "356",
    "MENDES": "357",
    "MIGUEL PEREIRA": "358",
    "MIRACEMA": "359",
    "NATIVIDADE-VARRE-SAI": "360",
    "NILÓPOLIS": "394",
    "NITERÓI": "397",
    "NOVA FRIBURGO": "395",
    "NOVA IGUAÇU-MESQUITA": "396",
    "PARACAMBI": "361",
    "PARAÍBA DO SUL": "362",
    "PARATY": "363",
    "PATY DO ALFERES": "2220",
    "PETRÓPOLIS": "398",
    "PINHEIRAL": "2331",
    "PIRAÍ": "364",
    "PORCIÚNCULA": "365",
    "PORTO REAL - QUATIS": "2218",
    "QUEIMADOS": "2209",
    "RESENDE": "399",
    "RIO BONITO": "366",
    "RIO CLARO": "367",
    "RIO DAS FLORES": "368",
    "RIO DAS OSTRAS": "2162",
    "SANTA MARIA MADALENA": "369",
    "SANTO ANTÔNIO DE PÁDUA-APERIBÉ": "370",
    "SÃO FIDELIS": "371",
    "SÃO FRANCISCO DO ITABAPOANA": "2194",
    "SÃO GONÇALO": "400",
    "SÃO JOÃO DA BARRA": "372",
    "SÃO JOÃO DE MERITI": "401",
    "SÃO JOSÉ DO VALE DO RIO PRETO": "2336",
    "SÃO PEDRO DA ALDEIA": "373",
    "SÃO SEBASTIÃO DO ALTO": "374",
    "SAPUCAIA": "375",
    "SAQUAREMA": "376",
    "SEROPÉDICA": "2345",
    "SILVA JARDIM": "377",
    "SUMIDOURO": "378",
    "TANGUÁ": "2380",
    "TERESÓPOLIS": "402",
    "TRAJANO DE MORAES": "379",
    "TRÊS RIO-AREAL-LEVY GASPARIANS": "403",
    "VALENÇA": "404",
    "VASSOURAS": "380",
    "VOLTA REDONDA": "405"
}

for nome_comarca, codigo in comarcas.items():
    url = base.format(comarca=codigo)
    response = requests.get(url)

    if response.status_code == 200:
        dados = response.json()
        registros = dados

        # Se quiser manter referência da comarca, adicione ao registro diretamente
        for registro in registros:
            registro["comarca"] = nome_comarca

        lista_final.extend(registros)
        print(f"Dados da comarca {nome_comarca} adicionados.")
    else:
        print(f"Falha na requisição da comarca {nome_comarca}: código {response.status_code}")

    time.sleep(1)

# Converter a lista JSON para DataFrame
df = pd.DataFrame(lista_final)

# Salvar DataFrame em arquivo CSV, sem índice
df.to_csv("dados.csv", index=False, encoding="utf-8-sig")

print("Arquivo CSV gerado com sucesso: dados.csv")