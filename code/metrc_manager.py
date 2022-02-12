# metric manager class
from logging import raiseExceptions
from tokenize import String
from splinter import Browser
import time
from selenium.webdriver.common.keys import Keys
import datetime
from verification import USERNAME, PW
import traceback
import pandas as pd

URL = "https://or.metrc.com/log-in?ReturnUrl=%2f"
class Metrc_Manager:
    username = USERNAME
    pw = PW
    today = datetime.date.strftime(datetime.date.today(), "%m/%d/%Y")
    employee_df = pd.read_csv("C:/Users/jsurg/code/TozMoz/employee_list.csv")
    reason_dict = {
        "API Adjustment Error" : 1,
        "API Conversion Error" : 2,
        "During Licence-to-Licence Transfer" : 3,
        "Entry Error" :4,
        "In-House Quality Control" : 5,
        "Moisture Loss/Gain" : 6,
        "Package Material" : 7,
        "Plant Death" : 8, 
        "Scale Variance" : 9,
        "Trade Sample" : 12,
        "Waste" : 13
    }
    unit_dict= {
        "Each" : 1,
        "Fluid Ounces" : 2,
        "Gallons" : 3,
        "Grams" : 4,
        "Kilograms" : 5,
        "Liters" : 6,
        "Milligrams" : 7,
        "Milliliters" : 8,
        "Ounces" : 9,
        "Pints" : 10,
        "Pounds" : 11,
        "Quarts" : 12
    }

    try:
        # connect browser variable to chrome using splinter
        executable_path = {'executable_path':'C:\\Users\\jsurg\\.wdm\\drivers\\chromedriver\\win32\\97.0.4692.71\\chromedriver.exe'}
        browser = Browser('chrome', **executable_path,headless=False)
    except Exception as e:
        print("chrome_browswer_main exceptions caught")
    # visit URL
    print("Opening Browser")
    browser.visit(URL)
    
    def __init__(self):
        print("Metric Manager Object Created")
        time.sleep(1)
        self.login()

        # initialize active button instance variable
        self.active_button = self.find_active()
        self.active_button.click()

    def login(self):
        print("Logging In")
        # Find the user and pass form and fill it
        self.browser.find_by_id('username').first.fill(self.username)
        self.browser.find_by_id('password').first.fill(self.pw)
        
        # find and click login button
        self.browser.find_by_id('login_button').first.click()
   
    def find_active(self):
        '''Returns splinter object that points to Active button'''
        # Find active button
        strip = self.browser.find_by_id("packages_tabstrip").first
        spans = strip.find_by_tag("span")
        return spans[1]
    
    def tag_search(self, to_search):
        '''
        Searches for a tag number. Returns a 1 if successful, 0 if not. 
        Leaves the browser at the page resulting from the search
        ''' 

        if(type(to_search) != str and type(to_search) != int):
            return 0 

        # find tag option --> check for active tag search button using try-except blocks
        try:
            tag_option = self.browser.find_by_css('a[class="k-header-column-menu k-state-active"]')[0]

            tag_option.click()
            time.sleep(1)

            # find filter tab option and click
            filter_button = self.browser.find_by_css('span[class="k-icon k-i-arrow-60-right k-menu-expand-arrow"]')[1]
            filter_button.click()
            time.sleep(1)
        
        except:
            tag_option = self.browser.find_by_css('a[class="k-header-column-menu"]')[0]

            tag_option.click()
            time.sleep(1)

            # find filter tab option and click
            filter_button = self.browser.find_by_css('span[class="k-icon k-i-arrow-60-right k-menu-expand-arrow"]')[1]
            filter_button.click()
            time.sleep(1)

        # find text box and fill it with passed tag
        self.browser.find_by_css('input[title="Filter Criteria"]').fill(to_search)

        # find submit button and click it
        submit = self.browser.find_by_css('button[class="k-button k-primary"]')
        submit.click()

        # wait before returning- allows page to load- serving up properly
        time.sleep(1)
        return 1

    def item_search(self, to_search):
        '''
        Searches for a tag number. Returns a 1 if successful, 0 if not. 
        Leaves the browser at the page resulting from the search
        ''' 

        if(type(to_search) != str):
            return 0 

        # find tag option --> check for active tag search button using try-except blocks
        try:
            tag_option = self.browser.find_by_css('a[class="k-header-column-menu"]')[5]

            tag_option.click()
            time.sleep(1)

            # find filter tab option and click
            filter_button = self.browser.find_by_css('span[class="k-icon k-i-arrow-60-right k-menu-expand-arrow"]')[1]
            filter_button.click()
            time.sleep(1)
        
        except:
            tag_option = self.browser.find_by_css('a[class="k-header-column-menu k-state-active"]')[0]

            tag_option.click()
            time.sleep(1)

            # find filter tab option and click
            filter_button = self.browser.find_by_css('span[class="k-icon k-i-arrow-60-right k-menu-expand-arrow"]')[1]
            filter_button.click()
            time.sleep(1)

        # find text box and fill it with passed tag
        self.browser.find_by_css('input[title="Filter Criteria"]').fill(to_search)

        # find submit button and click it
        submit = self.browser.find_by_css('button[class="k-button k-primary"]')
        submit.click()

        # wait before returning- allows page to load- serving up properly
        time.sleep(1)
        return 1

    def grab_info(self, history = False, lab_results = False):
        '''
        Returns a dictionary of info scraped from the current page, if found. returns None if not found.
        Leaves browser in same state as it was before the call.
        Keys: tag, source, loc, item, category, value, unit, prod batch, testing, pack date
        '''
        try:
            td_element = self.browser.find_by_tag('td') 
            value, unit = td_element[26].value.split(" ")
            info = {
                "tag" : td_element[1].value, 
                "source" : td_element[2].value, 
                "loc" : td_element[4].value, 
                "item" : td_element[6].value, 
                "category" : td_element[8].value, 
                "value" : value,
                "unit" : unit,
                "prod batch" : td_element[29].value, 
                "testing" : td_element[31].value, 
                "pack date" : td_element[43].value
                }
        except Exception as e: # tag does not exist
            print(f'Could not grab info, folloiwng exception raised:\n{e}')
            return None

    # asking for history or lab result
        if(history or lab_results):
            
            # click expander
            expander = self.browser.find_by_css('a[class="k-icon k-i-expand"]')
            expander.click()
            
            # history info
            if(history):
                self.browser.find_by_css('span[class="k-link"]')[7].click()
                ### to do: grab data
                
            if(lab_results):
                self.browser.find_by_css('span[class="k-link"]')[6].click()
                ### to do: grab data

        else:
            history_results = None
            lab_results = None
            
        # need to figure out how i want to return data if history/labs are requested
        return info   

    def new_package(self, tag, amount, loc, unit = "Grams", new_item = False, next_avail = True, prod_batch = False, finish = False):
        '''
        Creates a new package out of the tag passed, with the value of amount.
        (should) Return new tag number if successful, 0 otherwise;;; CURRENTLY RETURNS 1 IF SUCCESSFUL
        Asks user to confirm information before making package
        Leaves browswer at new package view, expecting user response
        '''

        # #### NEED TO ADD HERE: check the tag for validity (does first cell on page == tag passed). if not, return 0
        # if str(tag) not in self.grab_info()["tag"]:
        #      return 0
        
        try:
            # find tag and click
            self.browser.find_by_css('td[role="gridcell"]').first.click()

            # find new package button and click
            self.browser.find_by_css('button[class="btn shadow js-toolbarbutton-1"]').click()
            time.sleep(7)
        
            # helper function to decide functionality based off next_available parameter. Nested because it should not be used by other functions
            def next_available(value):
                '''
                Check truth value of parameter passed.
                True == grab next available tag
                False == let human decide next package number
                Returns 1 if successful, 0 if not
                '''
                try:
                    if value: # grab next available tag

                        # find tag search button and click
                        self.browser.find_by_css('button[class="js-tagsearch-0 btn js-tagsearch"]').click()

                        # wait one second
                        time.sleep(1)

                        # find next available tag and click
                        self.browser.find_by_text('CannabisPackage')[0].click()

                        # find select button and click
                        self.browser.find_by_css('button[id="entityselector-ok-button"]').click()
                        return 1

                    else:   # user manually specifies new tag to use. No action required
                        return 1
                except Exception as e:
                    raiseExceptions(e)

            ############ METRC IS SLOW RIGHT NOW SO ADDING SLEEP, REMOVE LATER
            time.sleep(5)

            # utilize next_available function to fill New Tag text box
            try: 
                next_available(next_avail)
            except Exception as e:  # New Package page did not load fast enough. Need to wait
                print("New Package page not loaded, trying to fill data again in 5 secs")
                time.sleep(5)
                next_available(next_avail)

            # NEEDED FUNCTIONALITY here: save new tag text to variable 
            #new_tag = self.browser.find_by_css('input[id="form-validation-field-0"]').text

            # Go to next element: location
            time.sleep(.1)
            active_web_element = self.browser.driver.switch_to.active_element
            time.sleep(.5)

            if next_avail:
                tab_press = range(3)
            else:
                tab_press = range(2)
            for i in tab_press:
                active_web_element.send_keys(Keys.TAB)
                active_web_element = self.browser.driver.switch_to.active_element
                time.sleep(.1)

            # Fill location
            active_web_element = self.browser.driver.switch_to.active_element
            time.sleep(.1)
            active_web_element.send_keys(loc)
            time.sleep(.1)

            # move active element down to item text by pressing ENTER key
            active_web_element.send_keys(Keys.ENTER)
            active_web_element.send_keys(Keys.ENTER)

            # use new_item to fill item text box
            if new_item:
                active_web_element = self.browser.driver.switch_to.active_element
                active_web_element.send_keys(new_item)
                active_web_element.send_keys(Keys.ENTER)

            else: 
                #default behavior: click same item button
                active_web_element = self.browser.driver.switch_to.active_element
                time.sleep(.1)
                for i in range(2):
                    active_web_element.send_keys(Keys.TAB)
                    active_web_element = self.browser.driver.switch_to.active_element
                    time.sleep(.1)
                # click same item
                active_web_element.click()

            # switch to next element: quantity unit
            for i in range(2):
                active_web_element.send_keys(Keys.TAB)
                active_web_element = self.browser.driver.switch_to.active_element
                time.sleep(.1)
            
            # fill unit measure--- default behavior is grams
            active_web_element.send_keys(unit)
            time.sleep(.1)
            active_web_element.send_keys(Keys.ENTER)
            # fill quantities
            for i in range(2):
                self.browser.find_by_css('input[class="validate[required,custom[number],min[0.0001]] span3 ng-pristine ng-valid-number ng-valid-min ng-invalid ng-invalid-required"]').fill(amount)

            try:
                # finish package
                if finish:
                    self.browser.find_by_css('input[name="model[0][Ingredients][0][FinishPackage]"]').click()
                    self.browser.find_by_css('input[name="model[0][Ingredients][0][FinishDate]"]').fill(today)

            except Exception as e:
                print("Could not finish tag. Following exception raised:\n", e)

            # fill date
            self.browser.find_by_css('input[name="model[0][ActualDate]"]').fill(today)
            
            # fill production batch section
            if prod_batch and isinstance(prod_batch, str): 
                # find and click production checkbox
                self.browser.find_by_css('input[name="model[0][IsProductionBatch]"]').click()

                # push active element to production text box
                active_web_element = self.browser.driver.switch_to.active_element
                active_web_element.send_keys(Keys.TAB)
                active_web_element = self.browser.driver.switch_to.active_element

                # fill production textbox
                active_web_element.send_keys(prod_batch)

            ##############
            # NEXT STEPS #
            ##############
            # to dos: conditional for items, conditional for today,
            # 
            input("Please Submit or Cancel New Package. Press Enter when done") 
            return 1 #NEEDS TO RETURN NEW TAG NUMBER
        
        except Exception as e:
            print(traceback.format_exc())
            return 0
        ######################################
        # new_package Needed Functionality:
            # needs functionality to allow for multiple 'right' packages into one 'left' package
            # needs to return new tag number
            # could ask user to confirm or cancel, but i think having the user actually click either is a better way of confirming accuracy
                # do we then need to wait for the user to make a selection (IE time.sleep)? In that case, user could just tell ...
                # ... the program continue when they are ready

    def adjust_package(self, tag, amount, unit, reason, date, confirm = False):
        '''
        Adjusts tag by amount passed in parameters
        '''
        # need to add tag check to make sure we're at the right page

        # find tag and click
        self.browser.find_by_css('td[role="gridcell"]').first.click()

        # find adjust button and click
        self.browser.find_by_css('button[class="btn shadow js-toolbarbutton-7"]').click()
        time.sleep(3)

        #################################
        #### Inside adjust menu  
        
        # try fill with date
        try:
            self.browser.find_by_css('input[class="validate[required,custom[dateFormat]] datepicker span3 ng-pristine ng-invalid ng-invalid-required"]').fill(date)
        except:
            # wait 5 seconds and try again
            time.sleep(5)
            self.browser.find_by_css('input[class="validate[required,custom[dateFormat]] datepicker span3 ng-pristine ng-invalid ng-invalid-required"]').fill(date)

        # find note element and fill with reason
        self.browser.find_by_css('input[name="model[0][ReasonNote]"]').fill(reason)

        # find New Quantity element and fill with amount
        self.browser.find_by_css('input[name="model[0][NewQuantity]"]').fill(amount) 
        
        if confirm:
            # have user confirm or cancel
            input("Please Submit or Cancel New Package. Press Enter when done")

        ##########################################
        # adjust_package Needed Functionality 
            # 1. fill reason list with appropriate reason
                ## find reason selector and click
                #sel_element = self.browser.find_by_css('select[class="validate[required] ng-pristine ng-valid ng-valid-required"]').click()
            # 2. do work with the unit parameter

    def waste_tag(self, tag, item):
        '''
        Returns 1 if successful, 0 if not
        '''
        # search tag and grab amount left in package
        self.tag_search(tag)
        time.sleep(1)
        weight_left = self.grab_info(tag)["value"]
    
        # call new_package(tag, weight_left, "Instake Area", new_item = item)
        self.new_package(tag, weight_left, "Intake Area", new_item = item, finish = True)
        return 1
        ############################
        # waste_tag needed functionality:
            # 1. tag_search needs to return the new tag, which waste_tag will return as well.

    def waste_tags_multiple_packages(self, tags, item):
        for tag in tags:
            self.waste_tag(tag, item)
            time.sleep(1)

    def zero_wasted(self, tag):

        # find tag and wait 1/2 second
        self.tag_search(tag)
        time.sleep(.5)

        # check for time package has been here 
        print(self.grab_info()["pack date"])

        # call self.adjust_package
        self.adjust_package(tag, 0, "Grams", "Package has been wasted after 3 days, following OLCC guidelines", today)

        # find proper reason selctor and click
        # self.browser.find_by_css('select[class="validate[required] ng-pristine ng-valid ng-valid-required"]').click()
        # self.browser.find_by_css('option[value="12"]').click()

        # finish tag and fill finish date
        # self.browser.find_by_css('input[name="model[0][Ingredients][0][FinishPackage]"]').click()
        # self.browser.find_by_css('input[name="model[0][Ingredients][0][FinishDate]"]').fill(today)

        # ask for confirmation
        input("Please Submit or Cancel New Package. Press Enter when done\n")

    def zero_multiple_waste(self, tag_list):
        skipped = []
        
        for tag in tag_list:
            try:
                self.zero_wasted(tag)
            except:
                print(f'skipped {tag}')
                skipped.append(tag)
            
            time.sleep(.75)

        return skipped
            
    def abc1_kill(self, items):
        grabbed_values = []
        for item in items:
            self.item_search(item)
            try:
                info = self.grab_info()
                print(info["tag"], info["item"])
                response = input("add to list? y/n ")
                if response =='y':
                    grabbed_values.append((info["tag"], info["item"]))

            except:
                print(f'{item} not found')
                continue

        return grabbed_values

    def qc_production_tag(self, tags, amount, loc):
        '''
        Produces new production batch tags of amount parameter quantity. Created out of tags parameter
        Leaves browser at new package view, expecting user response
        '''
        for i in range(len(tags)):
            self.tag_search(tags[i])
            prod_name = self.grab_info()["item"] + ' 220208'
            self.new_package(tags[i], amount, loc, prod_batch = prod_name)
        return 1

    def adjust_multi(self, tags, new_amount, unit, reason, reason_note):
        '''
        Adjusts multiple tags via one Adjust Submission. New Quantity = new_amount
        Leaves Metrc at packages view
        '''
        if unit not in self.unit_dict:
            raise ValueError("Invalid unit passed")
        
        if reason not in self.reason_dict:
            raise ValueError("Invalid reason passed")

        # find adjust button and click
        self.browser.find_by_css('button[class="btn shadow js-toolbarbutton-7"]').click()
        time.sleep(10)
        
        # create sufficient amount of package boxes
        self.browser.find_by_css('input[class= "span2 ng-pristine ng-valid ng-valid-number ng-valid-min"]').fill(len(tags)-1)
        time.sleep(1)
        self.browser.find_by_css('span[class="icon-plus"]').click()
        time.sleep(1)

        self.browser.find_by_css('select[class="span4 ng-pristine ng-valid"]').click()
        
        selector = self.unit_dict[unit]     # reference unit_dict to find number of DOWN ARROW keys to send
        for i in range(selector):
            active = self.browser.driver.switch_to.active_element
            active.send_keys(Keys.ARROW_DOWN)
        active.send_keys(Keys.ENTER)
        active.send_keys(Keys.TAB)
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.ENTER)

        # reason 
        self.browser.find_by_css('select[class="span5 ng-pristine ng-valid"]').click()
        active = self.browser.driver.switch_to.active_element

        selector = self.reason_dict[reason]     # reference reason_dict to find number of DOWN ARROW keys to send

        for i in range(selector): 
            active = self.browser.driver.switch_to.active_element
            active.send_keys(Keys.ARROW_DOWN)
        active.send_keys(Keys.ENTER)
        active.send_keys(Keys.TAB)
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.ENTER)

        # note
        self.browser.find_by_css('input[ng-model="template.ReasonNote"]').fill(reason_note)
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.TAB)
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.ENTER)

        # today's date
        self.browser.find_by_text('today').click()
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.TAB)
        active = self.browser.driver.switch_to.active_element
        active.send_keys(Keys.ENTER)
        for i in range(len(tags)):
            # package textbox
            self.browser.find_by_css(f'input[class="js-validated-element js-typeaheadsearchorder-{i} validate[required]"]').fill(tags[i])
            self.browser.driver.switch_to.active_element.send_keys(Keys.ENTER)
            
            # new quantity text box
            self.browser.find_by_css(f'input[name="model[{i}][NewQuantity]"]').fill(new_amount)

        # get user confirmation
        input("Please Submit or Cancel on Metrc")
        #####################################
        # adjust_multi needed functionality #
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ #
        # currently it only specifically does IHQC for reason, and grams for units-- need to make more dynamic
        # possibly add functionality for varying amounts
        # need functionality for finishing 
        return 1
         
###################################################################################################################################################
###################################################################################################################################################
###################################################################################################################################################

# test prep variables
today = datetime.date.strftime(datetime.date.today(), "%m/%d/%Y")
metrc_man = Metrc_Manager()
time.sleep(3)
print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')

# test class methods
# metrc_man.waste_tag(24665, "ABC1, Post-Extraction")
####################

abc1_items = [
    "Agent Orange Fresh Frozen",
 #x   "Angel Flower - Dream Queen",
    "Blue Magoo Fresh Frozen",
    "COOKIES AND CREAM",
    "Cookies and Cream Fresh Frozen",
  #  "dream queen",
  #x  "Dream Queen Fresh Frozen",
    "Fresh Frozen Agent Orange",
    "Fresh Frozen Blue Magoo",
  #x  "Fresh Frozen Dream Queen",
    "Fresh Frozen Gorilla Glue",
    "Fresh Frozen Jack the Ripper",
    "Fresh Frozen Jagermeister",
  #x  "Fresh Frozen Lemon Sour Diesel",
    "Fresh Frozen LSD",
    "Fresh Frozen Sauce",
    "Fresh Frozen Sherbert",
    "Fresh Frozen Sunset",
    "Fresh Frozen Sunset Sherbert",
   # "gorilla glue",    
   # "Gorilla Glue Fresh Frozen",
   # "mango",
   #---------------------
   # "MOB BOSS",
  #  "sherbert",
    "Skywalker Fresh Frozen",
   # "Sunset Sherbert Fresh Frozen",
    "Telos Agent Orange",
    "Telos Mob Boss"
    ]

tags = [24994, 24995, 24996]
metrc_man.adjust_multi(tags, 1, "Grams", "Waste", "a reason")


print("PROGRAM ENDING IN 15 secs")
time.sleep(15)
