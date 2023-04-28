import web_automation
import excel_automation
from web_automation import create_driver, login, capital_IQ, move_file
from excel_automation import mass_convert_xls_to_xlsm, mass_analysis


def main():
    username = input("Enter your Babson email -> ")
    password = input("Enter your password -> ")
    companies = input(
        "Enter companies of interest; separate each by a comma -> ")
    driver = create_driver()
    login(driver, "https://secure.signin.spglobal.com/sso/saml2/0oa1mqx8p77XSX10T1d8/app/spglobaliam_sp_1/exk1mregn1oWwP2NB1d8/sso/saml?RelayState=https://www.capitaliq.com/CIQDotNet/saml-sso.aspx", username, password)
    try:
        companies_list = capital_IQ(driver, companies)
        move_file(username, companies_list)
    except:
        raise ValueError(
            "[!] Incorrect company name or format; if entering multiple companies, please include commas between company names.")

    local_username = username.split("@")[0]
    folder_path = f"C:/Users/{local_username}/Documents/GitHub/Team-Project/excel_files"
    mass_convert_xls_to_xlsm(folder_path)
    mass_analysis(folder_path)
    print(f"-"*50, "\n Benchmarking was Successfully Executed \n", "-"*50)


if __name__ == '__main__':
    main()
