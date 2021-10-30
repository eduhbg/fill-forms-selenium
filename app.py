from selenium import webdriver
import time
import openpyxl

DRIVER = webdriver.Firefox()

def main():
    fill_inputs()

def fill_inputs():
    # File with information
    xlsx_file = 'finacial_sample.xlsx'
    wb_obj = openpyxl.load_workbook(xlsx_file)
    sheet = wb_obj.active
    
    # Define the start row
    row = 2

    # Run until finds a blank line
    while sheet['A' + str(row)].value != None:

        access_link()
        
        # Fill Segment
        DRIVER.find_element_by_xpath('/html/body/div/div[2]/form/div[2]/div/div[2]/div[1]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(sheet['A' + str(row)].value)

        # Fill Country
        DRIVER.find_element_by_xpath('/html/body/div/div[2]/form/div[2]/div/div[2]/div[2]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(sheet['B' + str(row)].value)

        # Fill Product
        DRIVER.find_element_by_xpath('/html/body/div/div[2]/form/div[2]/div/div[2]/div[3]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(sheet['C' + str(row)].value)

        # Fill Profit
        DRIVER.find_element_by_xpath('/html/body/div/div[2]/form/div[2]/div/div[2]/div[4]/div/div/div[2]/div/div[1]/div/div[1]/input').send_keys(sheet['L' + str(row)].value)
        
        # Get Date
        converted_date = date_to_str(sheet['M' + str(row)].value)

        # Fill Month
        DRIVER.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div/div[1]/div/div[2]/div[1]/div/div[1]/input').send_keys(converted_date[1])
        
        # Fill Day
        DRIVER.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div/div[3]/div/div[2]/div[1]/div/div[1]/input').send_keys(converted_date[0])
        
        # Fill Year
        DRIVER.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[2]/div/div[2]/div[5]/div/div/div[2]/div/div/div[5]/div/div[2]/div[1]/div/div[1]/input').send_keys(converted_date[2])
        
        # Send form
        DRIVER.find_element_by_xpath('/html/body/div[1]/div[2]/form/div[2]/div/div[3]/div[1]/div[1]/div/span').click()

        print(f'Line {row}')
        row += 1

        # Wait before initiate again
        time.sleep(1)

    print('Finished!')

def access_link():
    # Access the link through the webdriver 
    DRIVER.get('https://docs.google.com/forms/d/e/1FAIpQLSduYxCJcyKJc-j6chBfJBytqwFGyoJz498DmA7yLGWsPdrtrA/viewform?vc=0&c=0&w=1&flr=0')

def date_to_str(value):
    # Convert the datatime object to string
    date_str = value.strftime('%m/%d/%Y')
    date_spl = date_str.split(sep='/')
    return date_spl

if __name__ == '__main__':
    main()