import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from auth_data import login, password
from parcer_xlsx import get_data_xlsx


def ec_selenuim(driver):
    try:
        driver.get("https://ec.adm-nao.ru/auth/login-page")
        time.sleep(3)

        login_ec = driver.find_element(By.ID, "login")
        login_ec.send_keys(login)
        password_ec = driver.find_element(By.ID, "password")
        password_ec.send_keys(password)
        time.sleep(2)

        password_ec.send_keys(Keys.ENTER)
        time.sleep(18)

        driver.find_element(By.ID, "ext-gen12").click()
        driver.find_element(By.ID, "ext-comp-1057").click()
        driver.find_element(By.ID, "ext-comp-1059").click()
        time.sleep(3)

        input = driver.find_element(By.CSS_SELECTOR, "input.x-form-text.x-form-field")
        input.clear()
        input.send_keys("2023/2024")
        with open("name.txt", "r") as file_read:
            name = file_read.readlines()
            name = [l.rstrip("\n") for l in name]

        time.sleep(1)

        input.send_keys(Keys.ENTER)
        abbiturients = get_data_xlsx()
        time.sleep(5)

        for abbiturient in abbiturients:
            if abbiturient["FIO"] not in name:
                driver.find_element(By.CSS_SELECTOR, "button.x-btn-text.add_item").click()
                time.sleep(2)

                select_input = driver.find_element(By.CSS_SELECTOR,
                                                   "input.x-form-text.x-form-field.m3-form-invalid.x-trigger-noedit")
                select_input.click()
                time.sleep(1)

                select_input.send_keys(Keys.ARROW_DOWN)
                select_input.send_keys(Keys.ENTER)
                time.sleep(1)

                driver.find_element(By.XPATH,
                                    '//div[@class="x-window-mc x-panel-body-noheader"]/fieldset[3]/legend').click()
                time.sleep(1)

                last_name = driver.find_element(By.NAME, "last_name")
                first_name = driver.find_element(By.NAME, "first_name")
                middle_name = driver.find_element(By.NAME, "middle_name")
                date_of_birth = driver.find_element(By.NAME, "date_of_birth")

                last_name.clear()
                first_name.clear()
                middle_name.clear()
                date_of_birth.clear()

                FIO = abbiturient["FIO"].split()

                last_name.send_keys(FIO[0])
                first_name.send_keys(FIO[1])
                middle_name.send_keys(FIO[2])
                date_of_birth.send_keys(abbiturient["Birthday"])
                time.sleep(1)

                driver.find_element(By.XPATH,
                                    '//div[@class="x-window-mc x-panel-body-noheader"]/fieldset[7]/legend').click()
                select_input_1 = driver.find_element(By.CSS_SELECTOR,
                                                     "input.x-form-text.x-form-field.m3-form-invalid.x-trigger-noedit")
                select_input_1.click()
                select_input_1.send_keys(Keys.ENTER)
                average_mark = driver.find_element(By.CSS_SELECTOR,
                                                   "input.x-form-text.x-form-field.x-form-num-field.m3-form-invalid")
                average_mark.clear()
                average_mark.send_keys(round(abbiturient["Average mark"], 3))
                time.sleep(1)

                driver.find_element(By.XPATH,
                                    '//div[@class="x-window-mc x-panel-body-noheader"]/fieldset[8]/legend').click()
                select_input_2 = driver.find_element(By.CSS_SELECTOR,
                                                     'input.x-form-text.x-form-field.m3-form-invalid.x-trigger-noedit')
                select_input_2.click()
                select_input_2.send_keys(Keys.ENTER)
                select_input_3 = driver.find_element(By.ID, 'st_types')

                select_input_3.click()
                if abbiturient['Average mark'] >= 3.9:
                    time.sleep(2)
                    select_input_3.send_keys(Keys.ARROW_UP)
                    select_input_3.send_keys(Keys.ENTER)
                else:
                    time.sleep(2)
                    select_input_3.send_keys(Keys.ENTER)
                    select_input_3.send_keys(Keys.ARROW_DOWN)
                    select_input_3.send_keys(Keys.ARROW_DOWN)
                    time.sleep(1)
                    select_input_3.send_keys(Keys.ENTER)
                driver.find_element(By.NAME, "is_this_first_speciality").click()
                driver.find_element(By.NAME, "unit_lic_was_read").click()
                driver.find_element(By.NAME, "unit_lic_study_perm_read").click()
                driver.find_element(By.NAME, "unit_lic_accreditation_read").click()
                time.sleep(1)

                driver.find_element(By.XPATH, '//ul[@class="x-tab-strip x-tab-strip-top"]/li[2]').click()
                if abbiturient['INFO'] == 1:
                    driver.find_element(By.XPATH,
                                        '//div[@class="x-grid3-body"]/div[1]/table/tbody/tr[1]/td[1]/div[1]/div').click()
                if abbiturient['EKONOM'] == 1:
                    driver.find_element(By.XPATH,
                                        '//div[@class="x-grid3-body"]/div[4]/table/tbody/tr[1]/td[1]/div[1]/div').click()
                if abbiturient['POVAR'] == 1:
                    driver.find_element(By.XPATH,
                                        '//div[@class="x-grid3-body"]/div[2]/table/tbody/tr[1]/td[1]/div[1]/div').click()
                if abbiturient['EKOLOG'] == 1:
                    driver.find_element(By.XPATH,
                                        '//div[@class="x-grid3-body"]/div[3]/table/tbody/tr[1]/td[1]/div[1]/div').click()

                driver.find_element(By.XPATH,
                                    '//*[contains(text(), "%s" )]' % 'Сохранить').click()

                with open("name.txt", "a") as file:
                    file.write("\n"+abbiturient["FIO"])

                time.sleep(3)
            else:
                print("EST")

        time.sleep(5)
    except Exception as ex:
        print(ex)
    finally:
        driver.close()
        driver.quit()


def main():
    driver = webdriver.Firefox()
    ec_selenuim(driver=driver)


if __name__ == "__main__":
    main()