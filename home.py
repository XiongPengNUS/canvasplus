import streamlit as st
import pandas as pd
import requests

from canvasapi import Canvas
from IPython.core.display import display,HTML
from io import BytesIO


@st.cache
def get_student_id(course):

    return [p.user['id'] for p in course.get_enrollments() 
            if p.role == 'StudentEnrollment' and 
            p.user['integration_id'] is not None]


@st.cache
def get_user_profile(course):

    return [p.get_profile() for p in course.get_users()]


def to_excel(df, profiles):

    columns = df.columns
    if 'Avatar' in columns:
        index = list(columns).index('Avatar')
    else:
        index = None


    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  

    if index is not None:
        for i, profile in enumerate(profiles):
            image_url = profile['avatar_url']
            response = requests.get(image_url)
            worksheet.insert_image(i+1, index, image_url, 
                                   {'image_data': BytesIO(response.content),
                                    'x_scale': 0.2, 'y_scale': 0.2})
            worksheet.set_row(i+1, 80)

    writer.save()
    processed_data = output.getvalue()
    return processed_data


st.write("# Canvas Plus!")
st.write("This App supports extra features for Canvas.")

st.write('---')
token = st.text_input(label='Access token: ', key='token', value='')
url = 'https://canvas.nus.edu.sg/'
st.write('---')

if token != '':
    canvas = Canvas(url, token)
    courses = {course.name: course.id for course in canvas.get_courses()}
    course_name = st.selectbox('Courses:', courses, index=0)
    course = canvas.get_course(courses[course_name])

    tasks = ['Download Student List', 'Download Discussion Data']
    task = selected_task = st.selectbox('What to do:', tasks, index=0)

    st.write('---')

    if task == tasks[0]:
        # Students ID:
        st_id = get_student_id(course)

        # Selected student information
        info = {'Avatar': 'avatar_url', 
                'Student Number': 'integration_id', 
                'Email': 'primary_email'}
        info_columns = st.multiselect('Student Info.: ', info, info)
        
        # Selected group categories
        group_cats = {cat.name: cat for cat in course.get_group_categories()}
        group_cat_columns = st.multiselect('Group Categories: ', group_cats)
        selected_cats = [group_cats[c] for c in group_cat_columns]

        
        index = []
        students = {key: [] for key in ['Name'] + info_columns}

        profiles = get_user_profile(course)
        profiles = [p for p in profiles if p['id'] in st_id]
        for profile in profiles:
            index.append(profile['id'])
            students['Name'].append(profile['name'])

            for key in info_columns:
                if key == 'Avatar':
                    image_url = profile[info[key]]
                    students[key].append(f'<img src="{image_url}" width=100>')
                else:
                    students[key].append(profile[info[key]])

        df = pd.DataFrame(students, index=index)

        for cat in selected_cats:
            for g in cat.get_groups():
                p_id = [p.id for p in g.get_users()]
                df.loc[p_id, cat.name] = g.name

        # profiles = [p for p in profiles if p.id in st_id]
        df_xlsx = to_excel(df, profiles)
        st.download_button('Download', df_xlsx, file_name= 'students.xlsx')
        st.write(df.to_html(escape=False), unsafe_allow_html=True)
        

        
