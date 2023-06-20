import csv
import json

import requests
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.append(['University Name', 'College Board Code', 'Years', 'State', 'City', 'Public/Private', 'School Size',
           'Rural/Subrural', 'Average Tuition Fee', 'In State Tuition', 'Out of State Tuition',
           'Students Receiving Financial Aid', 'Application Types Accepted', 'Online Application Website',
           'GPA 4.0+', 'GPA 3.75+', 'GPA 3.50–3.74', 'GPA 3.25–3.49', 'GPA 3.00–3.24', 'GPA 2.50–2.99', 'GPA 2.00–2.49',
           'GPA Below 2.00', 'SAT total range lower bound', 'SAT total range higher bound', 'SAT math higher bound',
           'SAT math lower bound', 'SAT reading higher bound', 'SAT reading lower bound', 'ACT higher bound',
           'ACT lower bound', '	Acceptance rate', 'Total applicants number', 'Admitted number', 'Enrolled number',
           'High school GPA required for application', 'High school rank required for application',
           'College prep courses required for application', 'SAT/ACT scores required for application',
           'Recommendations required for application', 'Application fee', 'Regular decision date',
           'Graduation rate', 'Retention rate', 'Students to faculty ratio', 'Total undergrad students',
           'Total graduate students', 'Full time students', 'Part time students', 'Black or african american',
           'Asians', 'Hispanic or latino', 'Multiracial', 'Native Americans', 'Pacific Islanders', 'Unknown',
           'White', 'International (non-citizen)', 'Out of state residence',
           ])


with open("Ids.csv", 'r') as file:
    reader = csv.reader(file)
    for row in reader:
        # print(row[0])
        headers = {
            'authority': 'cs-search-api-prod.collegeplanning-prod.collegeboard.org',
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'en-GB,en;q=0.9,ar-SA;q=0.8,ar;q=0.7,en-US;q=0.6',
            'content-type': 'application/json',
            'origin': 'https://bigfuture.collegeboard.org',
            'referer': 'https://bigfuture.collegeboard.org/',
            'sec-ch-ua': '"Google Chrome";v="111", "Not(A:Brand";v="8", "Chromium";v="111"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"macOS"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36',
        }

        json_data = {
            'eventType': 'fetchByIds',
            'eventData': {
                'config': {},
                'criteria': {
                    'orgIds': [
                        f'{row[0]}',
                    ],
                    'fullPayload': 'full',
                    'rmsFilterInput': '',
                    'rmsFilterValue': '',
                },
            },
        }

        response = requests.post(
            'https://cs-search-api-prod.collegeplanning-prod.collegeboard.org/colleges',
            headers=headers,
            json=json_data,
        )
        jsn = json.loads(response.text)
        data = jsn['data']

        for each in data:
            name = each['name']
            print(f"Name:{name}")

            u_board_code = each['diCode']
            print(f"Board Code: {u_board_code}")

            year = each['schoolTypeByYears']
            print(f"schoolTypeByYears: {year}")

            rural = each['schoolSetting']
            print(f"schoolSetting:{rural}")

            public = each['schoolTypeByDesignation']
            print(f"schoolTypeByDesignation: {public}")

            school_size = each['schoolSize']
            print(f"School Size:{school_size}")

            state = each['stateName']
            print(f"state:{state}")

            city = each['city']
            print(f"city:{city}")

            net_price = each['averageNetPrice']
            net_price = f"${net_price}"
            print(f"netprice:{net_price}")

            in_station = each['inStateTuition']
            in_station = f"${in_station}"
            print(f"In station:{in_station}")

            out_station = each['outOfStateTuition']
            print(f"out station:{out_station}")

            student_aid = each['studentsReceivingAidPercent']
            student_aid = f"{student_aid}%"
            print(f"Student aid:{student_aid}")

            app_accepted = each['applicationsAccepted']
            if len(app_accepted) > 0:
                app_accepted = app_accepted[0]
            else:
                app_accepted = " "
            print(f"Application accepted:{app_accepted}")

            application_url = each['applicationSiteUrl']
            print(f"Application Url: {application_url}")

            #Admissions..............................................

            gpa1 = each['gpa400']
            gpa1 = f"{gpa1}%"
            print(f"{gpa1}")

            gpa2 = each['gpa375To399']
            gpa2 = f"{gpa2}%"
            print(gpa2)

            gpa3 = each['gpa350To374']
            gpa3 = f"{gpa3}%"
            print(gpa3)

            gpa4 = each['gpa325To349']
            gpa4 = f"{gpa4}%"
            print(gpa4)

            gpa5 = each['gpa300To324']
            gpa5 = f"{gpa5}%"
            print(gpa5)

            gpa6 = each['gpa250To299']
            gpa6 = f"{gpa6}%"
            print(gpa6)

            gpa7 = each['gpa200To249']
            gpa7 = f"{gpa7}%"
            print(gpa7)

            gpa8 = each['gpa100To199']
            gpa8 = f"{gpa8}%"
            print(gpa8)

            print("gpa ended.............................")

            sat_lower = each['satCompositeScore25thPercentile']
            print(sat_lower)

            sat_higher = each['satCompositeScore75thPercentile']
            print(sat_higher)

            sat_math_lower = each['rsatMathScore25thPercentile']
            print(sat_math_lower)

            sat_math_higher = each['rsatMathScore75thPercentile']
            print(sat_math_higher)

            sat_read_lower = each['rsatEbrwScore25thPercentile']
            print(sat_read_lower)

            sat_read_higher = each['rsatEbrwScore75thPercentile']
            print(sat_read_higher)

            act_lower = each['actCompositeScore25thPercentile']
            print(act_lower)

            act_higher = each['actCompositeScore75thPercentile']
            print(act_higher)

            acceptance_rate = each['acceptanceRate']
            acceptance_rate = f"{acceptance_rate}%"
            print(acceptance_rate)

            total_applicants = each['totalApplicants']
            print(total_applicants)

            admitted_num = each['admittedApplicants']
            print(admitted_num)

            enrolled_num = each['enrolledApplicants']
            print(enrolled_num)

            high_school_gpa = each['highSchoolGpa']
            print(high_school_gpa)

            high_school_rank = each['highSchoolRank']
            print(high_school_rank)

            college_prep_courses = each['prepCourses']
            print(college_prep_courses)

            sat_or_act_req = each['satOrAct']
            print(sat_or_act_req)

            recommendation_req = each['recommendations']
            print(recommendation_req)

            application_fee = each['applicationFeeAmount']
            print(application_fee)

            regular_decision_date = each['regularDecisionDate']
            print(regular_decision_date)

            #Academics..........................................

            graduation_rate = each['graduationRatePercent']
            graduation_rate = f"{graduation_rate}%"
            print(graduation_rate)

            retention_rate = each['sophomoreYearReturnPercent']
            retention_rate = f"{retention_rate}%"
            print(retention_rate)

            student_to_faculty = each['studentFacultyRatio']
            student_to_faculty = f"{student_to_faculty}:1"
            print(student_to_faculty)

            #CampusLife.......................................

            total_undergrads = each['totalUndergraduates']
            print(total_undergrads)

            total_grads = each['totalGraduates']
            print(total_grads)

            full_time_students = each['fullTimeEnrolled']
            print(full_time_students)

            part_time_students = each['partTimeEnrolled']
            print(part_time_students)

            black_or_african = each['africanAmericanPercent']
            black_or_african = f"{black_or_african}%"
            print(black_or_african)

            asians = each['asianPercent']
            asians = f"{asians}%"
            print(asians)

            hispanic = each['hispanicPercent']
            hispanic = f"{hispanic}%"
            print(hispanic)

            multi_racial = each['multiracialPercent']
            multi_racial = f"{multi_racial}%"
            print(multi_racial)

            native_american = each['nativeAmericanPercent']
            native_american = f"{native_american}%"
            print(native_american)

            pacific_islander = each ['pacificIslanderPercent']
            pacific_islander = f"{pacific_islander}%"
            print(pacific_islander)

            unknown = each['unknownPercent']
            unknown = f"{unknown}%"
            print(unknown)

            white = each['whitePercent']
            white = f"{white}%"
            print(white)

            international = each['internationalPercent']
            international = f"{international}%"
            print(international)

            out_of_state = each['outOfStatePercent']
            out_of_state = f"{out_of_state}%"
            print(out_of_state)

            #Saving Data Here................................

            ws.append([name, u_board_code, year, state, city, public, school_size,
                       rural, net_price, in_station, out_station,
                       student_aid, app_accepted, application_url,
                       gpa1, gpa2, gpa3, gpa4, gpa5, gpa6,
                       gpa7,
                       gpa8, sat_lower, sat_higher, sat_math_higher,
                       sat_math_lower, sat_read_higher, sat_read_lower, act_higher,
                       act_lower, acceptance_rate, total_applicants, admitted_num,
                       enrolled_num,
                       high_school_gpa, high_school_rank,
                       college_prep_courses, sat_or_act_req,
                       recommendation_req, application_fee, regular_decision_date,
                       graduation_rate, retention_rate, student_to_faculty, total_undergrads,
                       total_grads, full_time_students, part_time_students, black_or_african,
                       asians, hispanic, multi_racial, native_american, pacific_islander, unknown,
                       white, international, out_of_state,
                       ])

        wb.save("Data.xlsx")







