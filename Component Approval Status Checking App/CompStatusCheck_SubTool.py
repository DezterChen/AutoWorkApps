# -*- coding: utf-8 -*-
"""
Created on Tue Sep 26 19:22:11 2017

@author: Dezter
"""
def CSC_SubTool(PLM_ID,PLM_PW,Col_Info,Content):
    from selenium import webdriver
    import time
    #import requests
    #from bs4 import BeautifulSoup as bs
    #import sys
    """ Personal Info """
    
    """ "Login in" """
    driver = webdriver.Chrome('C:\\Anaconda3\\selenium\\webdriver\\chromedriver')
    driver.implicitly_wait(30)
    url = "http://pdmagile.quantatw.com/Agile/default/login-cms.jsp"
    
    driver.maximize_window()
    driver.get(url)
    driver.find_element_by_id("j_username").send_keys(PLM_ID)
    driver.find_element_by_id("j_password").send_keys(PLM_PW)
    driver.find_element_by_id("login").click()
    
    """ Close Extra Window """
    driver.switch_to.window(driver.window_handles[1])
    driver.close()
    """ Home page """
    driver.switch_to.window(driver.window_handles[0])
    driver.get("http://pdmagile.quantatw.com/Agile/PLMServlet?module=LoginHandler&opcode=forwardToMainMenu")
    
    """ Pre work """
    [Col_QPN,Col_MFC,Col_Aprv,Col_MD] = Col_Info
    
    """ Crawler """
    for i in range(len(Content)):
        """ Check Comp-Aprv """
        if Content[i][Col_Aprv].strip() == "Aprv done" and Content[i][Col_MD].strip() == "MD done":
            continue
        elif Content[i][Col_Aprv].strip() == "Aprv done" and Content[i][Col_MD].strip() != "MD done":
             Content[i][Col_MD] = "MD done"
        else:
            PN=Content[i][Col_QPN].strip()#Quanta P/N
            driver.find_element_by_name("QUICKSEARCH_STRING").clear()
            time.sleep(0.5)
            driver.find_element_by_name("QUICKSEARCH_STRING").send_keys(PN)
            time.sleep(0.5)
            driver.find_element_by_id("top_simpleSearch").click()
            time.sleep(2)
            if driver.find_element_by_class_name("header_wrapper").text == "Search Criteria":
                driver.find_element_by_link_text(PN).click()
                time.sleep(0.5)
                driver.find_element_by_link_text("Manufacturers").click()
            else:
                driver.find_element_by_link_text("Manufacturers").click()
            time.sleep(0.5)
            QPNInfo = driver.find_element_by_class_name("GMBodyMid").text
            MFC = Content[i][Col_MFC].strip() #MFC Code
            MPNInfo = QPNInfo[QPNInfo.find(MFC):QPNInfo.find("\n",QPNInfo.find(MFC),-1)]
            if "FR-Approve" in MPNInfo:
                CompStatus = "Aprv done"
            elif "NonApprove" in MPNInfo:
                driver.find_element_by_link_text("Changes").click()
                time.sleep(0.5)
                AprvApply = driver.find_elements_by_class_name("GMBodyMid")[0].text            
                if AprvApply[0:5].find("A") -2 == 0:
                    AprvApply = AprvApply.split( )
                    CompStatus = AprvApply[0]
                elif AprvApply[0:5].find("A") -2 != 0:
                    CompStatus = "Aprv wait"             
            
            """ Check MD release """        
            if Content[i][Col_MD].strip() == "MD done":
                continue
            else:                   
                MPN = MPNInfo.split( )[1] 
                driver.find_element_by_name("QUICKSEARCH_STRING").clear()
                time.sleep(0.5)
                driver.find_element_by_name("QUICKSEARCH_STRING").send_keys(MPN)
                time.sleep(0.5)
                driver.find_element_by_id("top_simpleSearch").click()
                time.sleep(2)
                if driver.find_element_by_class_name("header_wrapper").text.find("Search Criteria") == 0:
                    driver.find_element_by_link_text(MPN+" ("+MFC+")").click() 
                    time.sleep(0.5)
                    driver.find_element_by_link_text("Composition").click()
                    time.sleep(0.5)
                elif driver.find_element_by_class_name("header_wrapper").text.find("Search Results") == 0:
                    driver.find_element_by_link_text("Composition").click()
                    time.sleep(0.5)
                    
                driver.find_element_by_css_selector("option[title=\"Active * \"]").click()
                time.sleep(0.5)
                MD1 = driver.find_elements_by_class_name("GMBodyMid")[1].text
                driver.find_element_by_css_selector("option[title=\"Pending\"]").click()
                time.sleep(0.5)
                MD2 = driver.find_elements_by_class_name("GMBodyMid")[1].text
                
                if len(MD1) != 0:
                    MDStatus = "MD done"
                elif len(MD2) != 0:
                    MDStatus = MD2.split( )[0]
                elif len(MD2) == 0:
                    MDStatus = "MD wait"
                Content[i][Col_Aprv] = CompStatus#--> issue
                Content[i][Col_MD] = MDStatus
    driver.close()    
        #print(PN, MFC, CompStatus, MDStatus)
    data = Content
    return data
