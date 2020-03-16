from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import xlrd
from xlwt import Workbook
import os
#lecture de fichier excel
fichier ="test jean.xlsx"
wb = xlrd.open_workbook(fichier)
# feuilles dans le classeur
feuilles= wb.sheet_names()
# lecture des données dans la première feuille
sh = wb.sheet_by_name(feuilles[0])
# lecture par colonne
nom_cour = sh.col_values(0)
n_sequence = sh.col_values(1)
titre_sequence = sh.col_values(2)
n_sequence_video = sh.col_values(3)
link = sh.col_values(4)
#fin de lecture
cour_actuelle=""

br=webdriver.Chrome(executable_path=r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe")
mail="odoo@groupecerco.com"
mdp="C3rcOIA@@2019"


br.get("http://192.168.3.245:8069/web/login")

base=br.find_element_by_xpath('/html/body/div[1]/div/div/div[1]/div[55]/a')
base.click()

email=br.find_element_by_xpath('/html/body/div/main/div/div/div/div/div[2]/div/div/form/div[2]/input')
pasword=br.find_element_by_xpath('/html/body/div/main/div/div/div/div/div[2]/div/div/form/div[3]/input')
valid_btn=br.find_element_by_xpath('/html/body/div/main/div/div/div/div/div[2]/div/div/form/div[4]/button')
email.send_keys(mail)
pasword.send_keys(mdp)
valid_btn.submit()


cc=-1
ccx=0
balise=0
chapitre_en_cours=""
cccn=0
br.get("http://192.168.3.245:8069/web?#action=181&model=op.course&view_type=list&menu_id=149")
time.sleep(3)
while True:
    br.get("http://192.168.3.245:8069/web?#action=181&model=op.course&view_type=list&menu_id=149")
    time.sleep(3)
    br.refresh()
    time.sleep(5)
    try:
        input_chercher_cours=br.find_element_by_xpath('/html/body/div[1]/main/div[1]/div[1]/div/div/input')
    except:
        br.refresh()
        time.sleep(3)
        input_chercher_cours=br.find_element_by_xpath('/html/body/div[1]/main/div[1]/div[1]/div/div/input')
        print("exception 1 levé")
    
    time.sleep(3)
    print(nom_cour[cccn])
    input_chercher_cours.send_keys(nom_cour[cccn])
    input_chercher_cours.send_keys(Keys.ENTER)
    try:
        cours_vue=br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/table/tbody/tr[1]/td[2]')
        time.sleep(2)
    except:
        cours_vue=br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/table/tbody/tr[1]/td[2]')
        time.sleep(2)
        print("exception 2 levé")
    cours_vue.click()
    time.sleep(2)
    
    try:
        tgl_champ_mtrl= br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/div/div[2]/div[4]/ul/li[4]')
        time.sleep(1)
        tgl_champ_mtrl.click()
        time.sleep(1) 
    except:
        print("exception 3 levé")
        time.sleep(2)
        tgl_champ_mtrl= br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/div/div[2]/div[4]/ul/li[4]')
        time.sleep(2)
        tgl_champ_mtrl.click()
        time.sleep(1)
    try:
        
        cours_mat_edit_bnt=br.find_element_by_xpath('/html/body/div[1]/main/div[1]/div[2]/div/div/div[1]/button[1]')
        cours_mat_edit_bnt.click()
        time.sleep(1)
    except:
        print("exception 4 levé")
        time.sleep(2)
        cours_mat_edit_bnt=br.find_element_by_xpath('/html/body/div[1]/main/div[1]/div[2]/div/div/div[1]/button[1]')
        cours_mat_edit_bnt.click()
        time.sleep(1)
    try:    
        ajt_sqc_tab_1=br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/div/div[2]/div[4]/div/div[4]/div/div[2]/table/tbody/tr[1]/td/a')
        ajt_sqc_tab_1.click()
        time.sleep(2)
    except:
        print("exception 5 levé")
        time.sleep(2)
        ajt_sqc_tab_1=br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/div/div[2]/div[4]/div/div[4]/div/div[2]/table/tbody/tr[1]/td/a')
        ajt_sqc_tab_1.click()
        time.sleep(2)
    
    for ns in n_sequence:
        cccn+=1
        cc+=1
        ccx+=1
        balise+=1
        chapitre_en_cours=titre_sequence[cc]
        if chapitre_en_cours == titre_sequence[ccx]:
            try:
                input_sequence_n=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[1]/td[2]/input')
                input_sequence_title=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[2]/td[2]/input')
                input_sequence_n.clear()
                input_sequence_title.clear()
                input_sequence_n.send_keys(int(ns))
                input_sequence_title.send_keys(titre_sequence[cc])
            except:
                time.sleep(3)
                input_sequence_n=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[1]/td[2]/input')
                input_sequence_title=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[2]/td[2]/input')
                input_sequence_n.clear()
                input_sequence_title.clear()
                input_sequence_n.send_keys(int(ns))
                input_sequence_title.send_keys(titre_sequence[cc])
            
            
            chemin1="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td/a"
            chemin2="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td[1]/input"
            chemin3="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td[2]/div/div/input"
            try:
                ajouter_ligne=br.find_element_by_xpath(chemin1)
                ajouter_ligne.click()
                time.sleep(1)
            except:
                time.sleep(2)
                ajouter_ligne=br.find_element_by_xpath(chemin1)
                ajouter_ligne.click()
                time.sleep(1)
            try:
                sequenceN_input=br.find_element_by_xpath(chemin2)
                sequenceN_input.clear()
                sequenceN_input.send_keys(int(n_sequence_video[cc]))
            except:
                time.sleep(2)
                sequenceN_input=br.find_element_by_xpath(chemin2)
                sequenceN_input.clear()
                sequenceN_input.send_keys(int(n_sequence_video[cc]))
            
            try:
                dpr_material=br.find_element_by_xpath(chemin3)
                dpr_material.click()
                time.sleep(2)
            except:
                time.sleep(2)
                dpr_material=br.find_element_by_xpath(chemin3)
                dpr_material.click()
                time.sleep(2)
            try:    
                rch_link_more=br.find_element_by_xpath('/html/body/ul/li[8]/a')
                rch_link_more.click()
                time.sleep(2)
            except:
                time.sleep(2)
                rch_link_more=br.find_element_by_xpath('/html/body/ul/li[8]/a')
                rch_link_more.click()
                time.sleep(2)
                
            fitre_btn=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/button')
            fitre_btn.click()
            ajouter_filtre0=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/button')
            time.sleep(1)
            ajouter_filtre0.click()
            time.sleep(1)
            filtre_champ_1=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[2]/select')
            filtre_champ_1.click()
            time.sleep(1)
            choise_url=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[2]/select/option[13]')
            choise_url.click()
            filtre_champ_1.click()
            imput_url_mtr=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[3]/input')
            imput_url_mtr.send_keys(link[cc])
            btn_choise_url=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[3]/button[1]')
            btn_choise_url.click()
            time.sleep(2)
            link_trv=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[1]/div/input')
            link_trv.click()
            
            time.sleep(2)
            link_trv1=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[2]/div/table/tbody/tr[1]')
            link_trv1.click()
        else:
            input_sequence_n=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[1]/td[2]/input')
            input_sequence_title=br.find_element_by_xpath('/html/body/div[4]/div/div/main/div/div/table[1]/tbody/tr[2]/td[2]/input')
            input_sequence_n.clear()
            input_sequence_title.clear()
            input_sequence_n.send_keys(int(ns))
            
            input_sequence_title.send_keys(titre_sequence[cc])

            chemin1="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td/a"
            chemin2="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td[1]/input"
            chemin3="/html/body/div[4]/div/div/main/div/div/table[2]/tbody/tr[2]/td/div/div[2]/table/tbody/tr["+str(balise)+"]/td[2]/div/div/input"
            ajouter_ligne=br.find_element_by_xpath(chemin1)
            ajouter_ligne.click()
            time.sleep(1)
            sequenceN_input=br.find_element_by_xpath(chemin2)
            sequenceN_input.clear()
            sequenceN_input.send_keys(int(n_sequence_video[cc]))
            
            dpr_material=br.find_element_by_xpath(chemin3)
            dpr_material.click()
            time.sleep(2)
            rch_link_more=br.find_element_by_xpath('/html/body/ul/li[8]/a')
            rch_link_more.click()
            time.sleep(2)
            fitre_btn=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/button')
            fitre_btn.click()
            ajouter_filtre0=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/button')
            time.sleep(1)
            ajouter_filtre0.click()
            time.sleep(1)
            filtre_champ_1=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[2]/select')
            filtre_champ_1.click()
            time.sleep(1)
            choise_url=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[2]/select/option[13]')
            choise_url.click()
            filtre_champ_1.click()
            imput_url_mtr=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[2]/span[3]/input')
            imput_url_mtr.send_keys(link[cc])
            btn_choise_url=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[3]/div[1]/div/div[3]/button[1]')
            btn_choise_url.click()
            time.sleep(2)
            link_trv=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[1]/div[1]/div/input')
            link_trv.click()
            
            time.sleep(2)
            link_trv1=br.find_element_by_xpath('/html/body/div[6]/div/div/main/div[2]/div/table/tbody/tr[1]')
            link_trv1.click()
            time.sleep(2)
            
            if nom_cour[cc]==nom_cour[ccx]:
                save_creat=br.find_element_by_xpath('/html/body/div[4]/div/div/footer/button[2]')
                save_creat.click()
                time.sleep(2)
                balise=0
            else:
                save_cours=br.find_element_by_xpath('/html/body/div[4]/div/div/footer/button[1]')
                save_cours.click()
                time.sleep(2)
                save_matr=br.find_element_by_xpath('/html/body/div[1]/main/div[1]/div[2]/div/div/div[2]/button[1]')
                save_matr.click()
                balise=0
                time.sleep(1)
                mtr_confirm=br.find_element_by_xpath('/html/body/div[1]/main/div[2]/div/div/div/div[1]/div[1]/button[1]')
                mtr_confirm.click()
                cour_btn=br.find_element_by_xpath('/html/body/header/nav/ul[2]/li[2]')
                break
             
