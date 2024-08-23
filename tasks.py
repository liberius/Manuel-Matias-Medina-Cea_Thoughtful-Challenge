from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from robocorp.tasks import task
from RPA.Robocloud.Items import Items
import re
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import os

browser = Selenium()
excel = Files()
items = Items()

# Diccionarios amigables para los filtros
sort_options = {
    "Relevance": "0",
    "Newest": "3",
    "Oldest": "2"
}

category_options = {
    "Stories": "00000188-f942-d221-a78c-f9570e360000",
    "Subsections": "00000189-9323-db0a-a7f9-9b7fb64a0000",
    "Videos": "00000188-d597-dc35-ab8d-d7bf1ce10000"
}

def close_cookies_banner():
    try:
        browser.click_element("css=button[class*='accept-cookies']")
        print("Cookie banner closed.")
    except Exception as e:
        print("No cookie banner found or couldn't close it.", e)

def wait_for_clickable_and_click(locator):
    try:
        browser.wait_until_element_is_visible(locator, timeout=15)
        browser.click_element(locator)
        print("Clicked on element:", locator)
    except Exception as e:
        print("Could not click on element:", locator, e)

def click_with_javascript(locator):
    try:
        browser.execute_javascript(f"document.querySelector('{locator}').click()")
        print("Clicked on element using JavaScript:", locator)
    except Exception as e:
        print("Could not click on element using JavaScript:", locator, e)

def press_escape_key():
    try:
        browser.press_keys(None, Keys.ESCAPE)
        print("Pressed ESCAPE key.")
    except Exception as e:
        print("Could not press ESCAPE key.", e)

def ensure_clickable_and_click(locator):
    attempts = 0
    while attempts < 4:
        try:
            print(f"Attempt {attempts + 1}: Trying to click the element.")
            browser.click_element(locator)
            print("Clicked on element:", locator)
            break
        except Exception as e:
            print(f"Attempt {attempts + 1} failed. Reason: {e}")
            
            if attempts == 0:
                close_cookies_banner()
            elif attempts == 1:
                wait_for_clickable_and_click(locator)
            elif attempts == 2:
                click_with_javascript(locator)
            elif attempts == 3:
                press_escape_key()
            
            attempts += 1
            time.sleep(2) 
        else:
            print("Element clicked successfully.")
            break

@task
def main():
    # Cargar parámetros desde el Work Item
    search_term, sort_selection, category_selection = load_work_item()  
    
    # Abre el sitio web de AP News
    browser.open_available_browser("https://apnews.com/")
    
    # Ajusta el tamaño de la ventana (opcional)
    browser.set_window_size(1476, 598)
    
    # Asegura que el icono de búsqueda sea clickeable e intenta hacer clic
    ensure_clickable_and_click("css:.icon-magnify > use")
    
    # Intento de encontrar e interactuar con el campo de búsqueda
    search_field = None
    attempts = 0
    while attempts < 5:
        try:
            search_field = browser.execute_javascript("return document.querySelector('input[name=q]')")
            
            if search_field:
                browser.scroll_element_into_view(search_field)
                browser.input_text("name:q", search_term)
                browser.press_keys("name:q", Keys.RETURN)
                print("Search term entered.")
                break
            
        except Exception as e:
            print(f"Attempt {attempts + 1} failed to interact with the search field. Reason: {e}")
            attempts += 1
            time.sleep(2)
    
    if search_field is None:
        raise Exception("No se pudo interactuar con el campo de búsqueda después de varios intentos.")
    
    apply_filters(sort_selection, category_selection)
    
    news_data = extract_news_data(search_term)  # Aquí se pasa el argumento necesario
    
    save_news_to_excel(news_data)
    
    browser.close_all_browsers()

def load_work_item():
    items.get_input_work_item()
    search_term = items.get_work_item_variable("search_term")
    sort_selection = items.get_work_item_variable("sort_selection")
    category_selection = items.get_work_item_variable("category_selection")
    return search_term, sort_selection, category_selection

def apply_filters(sort_selection, category_selection):
    if sort_selection in sort_options:
        sort_value = sort_options[sort_selection]
        sort_select_xpath = f"//select[@name='s']/option[@value='{sort_value}']"
        
        try:
            browser.wait_until_element_is_visible(sort_select_xpath, timeout=10)
            browser.click_element(sort_select_xpath)
        except Exception as e:
            print(f"Failed to apply sort filter. Reason: {e}")
    
    for category in category_selection:
        if category in category_options:
            category_value = category_options[category]
            checkbox_xpath = f"//input[@value='{category_value}']"
            
            try:
                browser.wait_until_element_is_visible(checkbox_xpath, timeout=10)
                browser.click_element(checkbox_xpath)
            except Exception as e:
                print(f"Failed to apply category filter for {category}. Reason: {e}")
    
    browser.wait_until_element_is_not_visible("css:div.loading-spinner")

def extract_news_data(browser, search_term):
    news_items = []
    news_list_xpath = "//div[@class='PageList-items-item']"
    news_elements = browser.find_elements(news_list_xpath)

    print(f"Se encontraron {len(news_elements)} elementos de noticias.")

    for news in news_elements:
        try:
            title = news.find_element(By.XPATH, ".//span[@class='PagePromoContentIcons-text']").text
            description = news.find_element(By.XPATH, ".//div[@class='PagePromo-description']//span[@class='PagePromoContentIcons-text']").text
            date = news.find_element(By.XPATH, ".//div[@class='PagePromo-date']//span").text
            image_src = news.find_element(By.XPATH, ".//picture//img").get_attribute("src")
            
            print(f"Noticia encontrada: {title}")

            # Contar las ocurrencias de la frase de búsqueda en título y descripción
            search_count = (title.lower().count(search_term.lower()) + 
                            description.lower().count(search_term.lower()))

            # Verificar si contiene valores monetarios
            contains_money = bool(re.search(r'\$\d+(?:,\d+)*(?:\.\d+)?|\d+ dollars|\d+ USD', title + description))

            news_items.append({
                "title": title,
                "date": date,
                "description": description,
                "image_src": image_src,
                "search_count": search_count,
                "contains_money": contains_money
            })
        except Exception as e:
            print(f"Error al procesar la noticia: {e}")
    
    return news_items



def save_news_to_excel(news_items):
    if not news_items:
        print("No se encontraron datos para guardar en Excel.")
        return
    
    output_dir = "output"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Directorio '{output_dir}' creado.")
    
    try:
        excel.create_workbook(f"{output_dir}/news_data.xlsx")
        
        headers = ["title", "date", "description", "image_src", "search_count", "contains_money"]
        excel.append_rows_to_worksheet([headers])
        
        for item in news_items:
            row = [item["title"], item["date"], item["description"], item["image_src"], item["search_count"], item["contains_money"]]
            excel.append_rows_to_worksheet([row])
        
        excel.save_workbook()
        print(f"Archivo Excel guardado correctamente en '{output_dir}/news_data.xlsx'.")
    except Exception as e:
        print(f"Error al guardar el archivo Excel: {e}")

if __name__ == "__main__":
    main()
