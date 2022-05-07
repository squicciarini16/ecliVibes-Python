import PySimpleGUI as sg
import pandas as pd

#Theme
sg.theme('DarkGrey8')

#Connecting user data to excel file
EXCEL_FILE = (r'C:\Users\Nicholas\OneDrive\Desktop\ecliVibesPython\dataEntry.xlsx')
df = pd.read_excel(EXCEL_FILE)

############################################################################### FIRST WINDOW #################################################################################################################################


def make_win1():
    layout = [[sg.Text('Welcome to ECLI Vibes Survey!'), sg.Text('    ', k='-OUTPUT-')],

             [sg.Text('Enter Name', size=(15,1)), sg.InputText(key = 'Enter Name')],

             [sg.Text('Pronouns', size=(15,1)), sg.Combo(['He/Him', 'She/Her', 'They/Them', 'Other'], key = 'Pronouns')],

             [sg.Text('Language', size=(15,1)), 
                                        sg.Checkbox('English', key = 'English'),
                                        sg.Checkbox('Spanish', key = 'Spanish'), 
                                        sg.Checkbox('Other', key = 'Other')],
             [sg.Text('Basic Agency Information')],

             [sg.Text('Agency Name', size=(15,1)),sg.InputText(key = 'Agency Name')],

             [sg.Text('Agency Email', size=(15,1)),sg.InputText(key = 'Agency Email')],

             [sg.Text('Agency Website', size=(15,1)),sg.InputText(key = 'Agency Website')],

            [sg.Text('Agency Phone Number', size=(15,1)),sg.InputText(key = 'Agency Phone Number')],

            [sg.Text('Tell Us More!', size=(15,1)),sg.InputText(key = 'Tell Us More!')],

            [sg.Text("Our agency's work focuses on the following services", size=(40,1)), sg.Combo(['Trauma-Informed Services','Volutary Services','Survivor-Centered Services',
                    'Labor Trafficking Specific Services','Sex Trafficking Specific Services','None of the above'], key = "Our agency's work focuses on the following services")],

            [sg.Text('How does your agency meet these service standards?', size = (40,1)), sg.Combo(['Policy and Proceudres','Survivor leadership', 'Staff Trainings', 
            'Flexible and supportive work environment','None of the above'], key = 'How does your agency meet these service standards')],
            [sg.Text('Housing Information')],

            [sg.Text('Do you offer housing for human trafficking survivors?', size=(15,1)),
                                                                                sg.Checkbox('Yes', key = 'Yes'),
                                                                                sg.Checkbox('No', key = 'No'),
                                                                                sg.Checkbox('Unsure', key = 'Unsure')],

            [sg.Text('What is your housing funding source?', size=(40,1)), sg.Combo(['Housing and Urban Development','Office on Violence Against Women', 
            'Office for Victims of Crime', 'Community based',"I don't know our housing funding source"], key = 'What is your housing funding source?')],

            [sg.Text('What type of housing is offered for human trafficking surviviors?', size = (50,1)), sg.Combo(['Emergency Shelter', 'Transitional Housing', 'Shared Housing', 
            'Permanent Supportive Housing', 'Permanent Rapid Housing','None of the above'], key = 'What type of housing is offered for human trafficking surviviors?')],

            [sg.Text('If you chose emergency shelter, what type of shelter', size = (50,1)), 
                                                                                sg.Checkbox('DV Emergency Shelter', key = 'DV Emergency Shelter'),
                                                                                sg.Checkbox('DSS Shelter', key = 'DSS Shelter'),
                                                                                sg.Checkbox('Tier II Shelter', key = 'Tier II Shelter'),
                                                                                sg.Checkbox('Charity-Based Shelter', key= 'Charity-Based Shelter'),
                                                                                sg.Checkbox('None', key = 'None')],

            [sg.Text('What are the time limitations for housing options?', size = (50,1)), sg.Combo(['1-3 days', 'up to 30 days', 'up to 90 days',
            'up to 6 months','6 months to 1 year', 'more than 1 year', 'There is no limitation','None'], key = 'What are the time limitations for housing options?')],

            [sg.Text('What area do you serve', size = (50,1)), 
                                                sg.Checkbox('Suffolk County', key = 'Suffolk County'), 
                                                sg.Checkbox('Nassau County', key = 'Nassau County'), 
                                                sg.Checkbox('All of Long Island', key = 'All of Long Island'),
                                                sg.Checkbox('None', key= 'None')],

                    [sg.Button('Launch 2nd Window'), sg.Submit(), sg.Button('Exit')]]
    return sg.Window('ECLI Vibes', layout, location=(800,600), finalize=True)

    
############################################################################### SECOND WINDOW #################################################################################################################################

def make_win2():
    layout = [[sg.Text('Next Section')],
    [sg.Text('Housing Eligibility Criteria')],
    [sg.Text('What age is required to be eligible for independent housing?', size = (50,1)), 
                                                                            sg.Checkbox('Under 12', key = 'Under 12'),
                                                                            sg.Checkbox('12-15', key = '12-15'),
                                                                            sg.Checkbox('15-18', key = '15-18'),
                                                                            sg.Checkbox('18-23', key = '18-23'),
                                                                            sg.Checkbox('23+', key = '23+'),
                                                                            sg.Checkbox('No age requirement', key = 'No age requirement')],

    [sg.Text('Housing is available for', size = (15,1)), sg.Combo(['Single identifying female','Single identifying male','Households without children',
    'Households with children','Housholds with children only under 18','Pregnant','None of the above'], key = 'Housing is available for')],
                                        
    [sg.Text('Housing is targeted for individuals with the following immigration status', size = (50,1)),
                                                                                        sg.Checkbox('US Citizen', key = 'US Citizen'),
                                                                                        sg.Checkbox('Documented foreign national', key = 'Documented foreign national'),
                                                                                        sg.Checkbox('Undocumented foreign national', key = 'Undocumented foreign national'),
                                                                                        sg.Checkbox('None of the above', key = 'None of the above')],
    [sg.Text('Is an Axis 1 diagnosis required to quality for housing?', size = (15,1)),
                                                                        sg.Checkbox('Yes', key = 'Yes'),
                                                                        sg.Checkbox('No', key = 'No'),
                                                                        sg.Checkbox('Unsure', key = 'Unsure'),
                                                                        sg.Checkbox('Not applicable', key = 'Not applicable')],
    [sg.Text('With regards to substance abuse, individuals', size = (15,1)), sg.Combo(['Must be sober', 'There is no barrier for substance abuse', 'Other'], key = 'With regards to substance abuse, individuals')],

    [sg.Text('Inclusive Housing Information')],

    [sg.Text('Which population do you target?', key = (15,1)), sg.Combo(['LGBTQ+', 'Youth', 'Black Indigenous People Of Color', 
    'Asian/Pacific Islander or Asian Americans and Pacific Islanders', 'Individuals with disabilities','None of the above'], key = 'Which population do you target?')],

    [sg.Text('What language access do you offer?', size = (15,1)), 
                                                    sg.Checkbox('English only', key = 'English only'),
                                                    sg.Checkbox('Multi-Language Speaking Providers', key = 'Multi-Language Speaking Providers'),
                                                    sg.Checkbox('Language Access Hotline', key = 'Language Access Hotline'),
                                                    sg.Checkbox('None', key = 'None')],
    
    [sg.Text('If you chose multi-language speaking providers, what languages do your service providers speak?', size = (50,1)), 
    sg.InputText(key = 'If you chose multi-language speaking providers, what languages do your service providers speak?')],

    [sg.Text('Is your agency Americans with Disabilities Act (ADA) compliant?', size = (15,1)), 
                                                                                    sg.Checkbox('Yes', key = 'Yes'),
                                                                                    sg.Checkbox('No', key = 'No'),
                                                                                    sg.Checkbox('Unsure', key = 'Unsure')],
     [sg.Text('What are your cultural competency policies and guidelines? Please explain', size = (50,1)), 
    sg.InputText(key = 'What are your cultural competency policies and guidelines? Please explain')],            

    [sg.Text('Barriers to Housing')],
    ############# MUST CHANGE CHECKBOX TO COMBO BOX
    [sg.Text('If you have ever been denied housing to a victim of human trafficking, what were the reasons?', size = (70,1)), 
                                                                                                                sg.Checkbox('Substance use', key = 'Substance use'),
                                                                                                                sg.Checkbox('Gang involvment', key = 'Gang involvment'),
                                                                                                                sg.Checkbox('Emergency housing not an option', key = 'Emergency housing not an option'),
                                                                                                                sg.Checkbox('Immigration issues', key = 'Immigration issues'),
                                                                                                                sg.Checkbox('Mental Health', key = 'Mental Health'),
                                                                                                                sg.Checkbox('Past criminal history', key = 'Past criminal history'),
                                                                                                                sg.Checkbox('Waitlist barrier', key = 'Waitlist barrier'),
                                                                                                                sg.Checkbox('None of the above', key = 'None of the above')],
    [sg.Text('If you chose any of the above options, please explain', size = (15,1)), 
    sg.InputText(key = 'If you chose any of the above options, please explain')],

    [sg.Text('Do you require client COVID-19 Testing', size = (15,1)), 
                                                        sg.Checkbox('Yes', key = 'Yes'),
                                                        sg.Checkbox('No', key = 'No'),
                                                        sg.Checkbox('Unsure', key = 'Unsure')],

    [sg.Text('Do you have any COVID-19 accommodations or restrictions? Please explain', size = (15,1)), 
    sg.InputText(key = 'Do you have any COVID-19 accommodations or restrictions? Please explain')],

    [sg.Text('Due to COVID-19, how are services being provided?', size = (15,1)),
                                                        sg.Checkbox('In-person', key = 'In-person'),
                                                        sg.Checkbox('Hybrid', key = 'Hybrid'),
                                                        sg.Checkbox('Online', key = 'Online'),
                                                        sg.Checkbox('None of the above', key = 'None of the above')],

    [sg.Text('Screening Process')],
    ############ CHANGE FROM CHECKBOX TO COMBO
    [sg.Text('How are survivors able to access services?', size = (15,1)),
                                                            sg.Checkbox('Walk-in', key='Walk-in'),
                                                            sg.Checkbox('Call hotline', key = 'Call hotline'),
                                                            sg.Checkbox('In-person intake', key= 'In-person intake'),
                                                            sg.Checkbox('Phone interview', key = 'Phone interview'),
                                                            sg.Checkbox('Outreach workers', key = 'Outreach workers'),
                                                            sg.Checkbox('None of the above', key = 'None of the above')],

    [sg.Text('What are your screening procedures? Please describe', size = (15,1)), 
    sg.InputText(key = 'What are your screening procedures? Please describe')], 

    [sg.Text('If you have a hotline, how is the hotline accessed?', size = (15,1)),
                                                                    sg.Checkbox('Survivor calls on their own', key = 'Survivor calls on their own'),
                                                                    sg.Checkbox('Survivor can call with a case worker', key = 'Survivor can call with a case worker'),
                                                                    sg.Checkbox('Case worker can call on behalf of the survivor', key = 'Case worker can call on behalf of the survivor'),
                                                                    sg.Checkbox('None of the above', key = 'None of the above')],

    [sg.Text('Setting Layout and Services')],
    ##### CHANGE FROM CHECKBOX TO COMBO
    [sg.Text('Please describe the housing setting', size = (15,1)), 
                                                    sg.Checkbox('Shared room', key = 'Shared room'),
                                                    sg.Checkbox('Private room in shared setting', key = 'Private room in shared setting'),
                                                    sg.Checkbox('Private setting', key = 'Private setting'),
                                                    sg.Checkbox('Accessibility', key = 'Accessibility'),
                                                    sg.Checkbox('Americans with Disabilities Act', key = 'Americans with Disabilities Act'),
                                                    sg.Checkbox('None of the above', key = 'None of the above')],
    ########## CHANGE FROM CHECKBOX TO COMBO
    [sg.Text('What supportive services are provided?', size = (15,1)), sg.Combo('Food onsite','Case management','Legal/Advocay services',)
                                                       
                                                       
                                                        sg.Checkbox('Employment support services', key = 'Employment support services'),
                                                        sg.Checkbox('Transportation', key = 'Transportation'),
                                                        sg.Checkbox('Language Access/Interpretation Services', key = 'Language Access/Interpretation Services'),
                                                        sg.Checkbox('Safety Planning', key = 'Safety Planning'),
                                                        sg.Checkbox('SOAR Certified staff', key = 'SOAR Certified staff'),
                                                        sg.Checkbox('Support with accessing benefits', key = 'Support with accessing benefits'),
                                                        sg.Checkbox('None of the above', key = 'None of the above')],
    
    [sg.Text('Is there any aspect of the housing setting that is unique? Please describe: EG- therapeutic horses, painting classes yoga onsite, etc...', size = (70,1)), 
    sg.InputText(key = 'Is there any aspect of the housing setting that is unique? Please describe: EG- therapeutic horses, painting classes yoga onsite, etc...')], 

    [sg.Text('Agency Rules')],

    [sg.Text('Do you have a program overview or is there an application of your program?', size = (15,1)),
                                                                                            sg.Checkbox('Yes (If yes, please email it to Molly England menengland@empowerli.org', key = 'Yes (If yes, please email it to Molly England menengland@empowerli.org'),
                                                                                            sg.Checkbox('No', key = 'No'),
                                                                                            sg.Checkbox('Unknown', key = 'Unknown')],

    [sg.Text('How often would you participate in ongoing human trafficking housing resource meetings to share concerns and updates?', size = (70,1)), sg.Combo(['Monthly', 'Twice Annualy', 'Quarterly', 'Annually',"I'm not interested in participating in any human trafficking housing resource meetings", 'None'], 
    key = 'How often would you participate in ongoing human trafficking housing resource meetings to share concerns and updates?' )],

    [sg.Text('Do you have any questions that you feel we should add to this survey? Please share any issues and concerns that you have. What barriers or gaps in service do you experience?', size = (100,1)), 
    sg.InputText(key = 'Do you have any questions that you feel we should add to this survey? Please share any issues and concerns that you have. What barriers or gaps in service do you experience?')], 
                                                                                                                                                    
              
            #   [sg.Input(key='-IN-', enable_events=True)],
            #   [sg.Text(size=(25,1), k='-OUTPUT-')],
              [sg.Submit(), sg.Button('Exit')]]
    return sg.Window('Window: 2', layout, finalize=True)





def main():
    window1, window2 = make_win1(), None        # start off with 1 window open

    while True:     
        window, event, values = sg.read_all_windows()
        if event == sg.WIN_CLOSED or event == 'Exit':
            window.close()
            if window == window2:       # if closing win 2, mark as closed
                window2 = None         
            elif window == window1:     # if closing win 1, exit program
                break
        elif event == 'Launch 2nd Window' and not window2:
            window2 = make_win2()
        elif event == 'Submit':
            df = df.append(values, ignore_index = True)
            df.to_excel(EXCEL_FILE, index= False)
            sg.popup('Survey Submitted!')
            
        window.close()
    
if __name__ == '__main__':
    main()


