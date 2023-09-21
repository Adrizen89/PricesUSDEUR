import requests
from bs4 import BeautifulSoup
from datetime import datetime


def ZLME_scrap(start_date=None, end_date=None):
    try:
        response = requests.get('https://www.westmetall.com/en/markdaten.php?action=table&field=Euro_MTLE', verify=False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        table = soup.find("table")
        rows = soup.find_all("tr")[1:]  # Ignorer l'en-tête du tableau

        months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }

        if start_date:
            start_date_str = start_date.strftime("%d.%m.%Y")
        if end_date:
            end_date_str = end_date.strftime("%d.%m.%Y")

        data_list = []
        for row in rows:  # Boucle à travers toutes les lignes sauf l'en-tête
            columns = row.find_all("td")
            
            # Vérifier qu'il y a suffisamment de colonnes avant de continuer
            if len(columns) < 2:
                continue
            
            # Récupérer et formater la date
            date_data_raw = columns[0].text.strip()
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data_str = f"{day}.{month_num}.{year}"

            # Convertir la date en objet datetime pour la comparaison
            date_data = datetime.strptime(date_data_str, "%d.%m.%Y").date()
            
            # Récupérer et formater la donnée
            data = columns[1].text.strip()
            formatted_data = data.replace('.', ',')

            # Si aucune plage de dates n'est spécifiée, retourner seulement la valeur du jour
            if start_date is None and end_date is None:
                return [(date_data_str, formatted_data)]
            
            # Si start_date et end_date sont les mêmes, renvoyer seulement la donnée du jour
            if start_date == end_date:
                return [(date_data_str, formatted_data)]
            
            # Vérifier si la date est dans la plage de dates spécifiée
            if start_date and date_data < start_date:
                continue
            if end_date and date_data > end_date:
                continue
            
            data_list.append((date_data_str, formatted_data))
            if start_date and end_date:
            # Inverser la liste pour afficher du plus ancien au plus récent
                data_list.reverse()
        
        return data_list
    except requests.RequestException as e:
        return f"Erreur de requête : {e}"
    except Exception as e:
        return f"Erreur : {e}"





    


def EURX_scrap(start_date = None, end_date = None):
    try:
        response = requests.get('https://www.westmetall.com/en/markdaten.php?action=table&field=Euro_EZB', verify= False)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        table = soup.find("table")
        rows = soup.find_all("tr")[1:]  # Ignorer l'en-tête du tableau

        months = {
            'January': '01', 'February': '02', 'March': '03', 'April': '04', 'May': '05', 
            'June': '06', 'July': '07', 'August': '08', 'September': '09', 'October': '10',
            'November': '11', 'December': '12'
        }

        if start_date:
            start_date_str = start_date.strftime("%d.%m.%Y")
        if end_date:
            end_date_str = end_date.strftime("%d.%m.%Y")

        data_list = []
        for row in rows:  # Boucle à travers toutes les lignes sauf l'en-tête
            columns = row.find_all("td")
            
            # Vérifier qu'il y a suffisamment de colonnes avant de continuer
            if len(columns) < 2:
                continue
            
            # Récupérer et formater la date
            date_data_raw = columns[0].text.strip()
            day, month_name, year = date_data_raw.replace('.', '').split()
            month_num = months.get(month_name, '00')
            date_data_str = f"{day}.{month_num}.{year}"

            # Convertir la date en objet datetime pour la comparaison
            date_data = datetime.strptime(date_data_str, "%d.%m.%Y").date()
            
            # Récupérer et formater la donnée
            data = columns[1].text.strip()
            formatted_data = data.replace('.', ',')

            # Si aucune plage de dates n'est spécifiée, retourner seulement la valeur du jour
            if start_date is None and end_date is None:
                return [(date_data_str, formatted_data)]
            
            # Si start_date et end_date sont les mêmes, renvoyer seulement la donnée du jour
            if start_date == end_date:
                return [(date_data_str, formatted_data)]
            
            # Vérifier si la date est dans la plage de dates spécifiée
            if start_date and date_data < start_date:
                continue
            if end_date and date_data > end_date:
                continue
            
            data_list.append((date_data_str, formatted_data))

            if start_date and end_date:
            # Inverser la liste pour afficher du plus ancien au plus récent
                data_list.reverse()
        
        return data_list
    except requests.RequestException as e:
        return f"Erreur de requête : {e}"
    except Exception as e:
        return f"Erreur : {e}"



