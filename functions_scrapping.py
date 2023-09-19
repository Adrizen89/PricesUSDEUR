import requests
from bs4 import BeautifulSoup


def ZLME_scrap():
    try:
        response = requests.get('https://www.westmetall.com/en/markdaten.php?action=table&field=Euro_MTLE', verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        table = soup.find("table")
        rows = soup.find_all("tr")
        second_row = rows[1]

        
        # Trouver la quatrième colonne de la table dans la deuxième ligne
        columns = second_row.find_all("td")
        fourth_column = columns[1]
        # Récupérer la date de la première colonne
        first_column = columns[0]
        date_data_raw = first_column.text.strip()
        # Conversion de la date du format "05. September 2023" à "05/09/2023"
        months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }

        day, month_name, year = date_data_raw.replace('.', '').split()
        month_num = months.get(month_name, '00')  # Si le mois n'est pas trouvé, '00' est utilisé par défaut
        date_data = f"{day}.{month_num}.{year}"


        # Extraire le texte de la quatrième colonne
        data = fourth_column.text.strip()
        formatted_data = data.replace('.', ',')

        return date_data, formatted_data
    except requests.RequestException as e:
        return f"Erreur de requête : {e}"
    except Exception as e:
        return f"Erreur : {e}"

def EURX_scrap():
    try:
        response = requests.get('https://www.westmetall.com/en/markdaten.php?action=table&field=Euro_EZB', verify= False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        
        table = soup.find("table")
        rows = soup.find_all("tr")
        second_row = rows[1]

        
        # Trouver la quatrième colonne de la table dans la deuxième ligne
        columns = second_row.find_all("td")
        fourth_column = columns[1]
        # Récupérer la date de la première colonne
        first_column = columns[0]
        date_data_raw = first_column.text.strip()
        # Conversion de la date du format "05. September 2023" à "05/09/2023"
        months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }

        day, month_name, year = date_data_raw.replace('.', '').split()
        month_num = months.get(month_name, '00')  # Si le mois n'est pas trouvé, '00' est utilisé par défaut
        date_data = f"{day}.{month_num}.{year}"


        # Extraire le texte de la quatrième colonne
        data = fourth_column.text.strip()
        formatted_data = data.replace('.', ',')

        return date_data, formatted_data
    except requests.RequestException as e:
        return f"Erreur de requête : {e}"
    except Exception as e:
        return f"Erreur : {e}"