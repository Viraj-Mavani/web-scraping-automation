import requests

url = 'https://ica.justice.gov.il/GenericCorporarionInfo/SearchGenericCorporation'
payload = {
    "UnitsType": 8,
    "CorporationName": 'aaa'
}

headers = {
    'Content-Type': 'application/json',  # Add this header if required
}

try:
    response = requests.post(url, json=payload, headers=headers, timeout=200)
    response.raise_for_status()  # Raise an exception if the request was not successful

    json_data = response.json()
    print(json_data)
except requests.exceptions.RequestException as e:
    print(f"Error: {e}")