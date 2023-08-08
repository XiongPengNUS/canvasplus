import streamlit as st
import pandas as pd
import requests
import math

from canvasapi import Canvas
from canvasapi.exceptions import InvalidAccessToken
from IPython.core.display import HTML
from io import BytesIO
from PIL import Image

@st.cache_resource
def get_roles(token, course_name):

    url = 'https://canvas.nus.edu.sg/'
    canvas = Canvas(url, token)
    courses = {course.name: course.id for course in canvas.get_courses()}
    course = canvas.get_course(courses[course_name])

    return {p.user['id']: p.role 
            for p in course.get_enrollments()}


@st.cache_resource
def get_user_profile(token, course_name, roles, selected_cat):

    url = 'https://canvas.nus.edu.sg/'
    canvas = Canvas(url, token)
    courses = {course.name: course.id for course in canvas.get_courses()}
    course = canvas.get_course(courses[course_name])

    profiles = []
    
    st.write("Retrieving students' profiles")
    profile_bar = st.progress(0)

    if selected_cat == '':
        all_users = list(course.get_users())
    else:
        group_cats = {cat.name: cat for cat in course.get_group_categories()}
        # cats = [group_cats[c] for c in selected_cats.split('+++')]
        all_users = []
        for g in group_cats[selected_cat].get_groups():
            all_users.extend(g.get_users())

    n = len(all_users)
    for i, p in enumerate(all_users):
        if roles[p.id] == 'StudentEnrollment' and p.name != 'Test student':
            profiles.append(p.get_profile())

        profile_bar.progress((i+1) / n)

    return profiles


@st.cache_resource
def get_group_idx(token, cours_name, cat_columns):

    url = 'https://canvas.nus.edu.sg/'
    canvas = Canvas(url, token)
    courses = {course.name: course.id for course in canvas.get_courses()}
    course = canvas.get_course(courses[course_name])

    group_cats = {cat.name: cat for cat in course.get_group_categories()}
    cats = cat_columns.split('+++')
    cat_dict = {}
    for cat in cats:
        group_dict = {}
        for g in group_cats[cat].get_groups():
            user_id = [p.id for p in g.get_users()]
            group_dict[g.name] = user_id
        cat_dict[cat] = group_dict
    
    return cat_dict


@st.cache_resource
def gen_preview_table(token, course_name, selected_cat, info_columns, info, cat_columns):

    # User roles:
    roles = get_roles(token, course_name)
    
    index = []
    students = {key: [] for key in ['Name'] + info_columns}

    profiles = get_user_profile(token, course_name, roles, selected_cat)
    for i, profile in enumerate(profiles):
        index.append(profile['id'])
        students['Name'].append(profile['name'])

        for key in info_columns:
            if key == 'Avatar':
                image_url = profile[info[key]]
                students[key].append(f'<img src="{image_url}" width=100>')
            else:
                students[key].append(profile[info[key]])
            
    df = pd.DataFrame(students, index=index)
    
    if cat_columns:
        cat_dict = get_group_idx(token, course_name, cat_columns)
        for cat_name, cat in cat_dict.items():
            for g_name, g_idx in cat.items():
                row_idx = list(set(g_idx).intersection(set(df.index)))
                df.loc[row_idx, cat_name] = g_name

    return  df, profiles


@st.cache_resource
def to_excel(token, course_name, selected_cat, info_columns, info, cat_columns):

    df, profiles = gen_preview_table(token, course_name, selected_cat, 
                                     info_columns, info, cat_columns)

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
        st.write('Generating avatars')
        data_bar = st.progress(0)
        n = len(profiles)
        for i, profile in enumerate(profiles):
            image_url = profile['avatar_url']
            response = requests.get(image_url)

            img = Image.open(BytesIO(response.content))
            # img = img.resize([math.ceil(d) for d in img.info['dpi']])
            byteIO = BytesIO()
            img.save(byteIO, format='PNG')
            worksheet.insert_image(i+1, index, image_url, 
                                   {'image_data': byteIO,
                                    'x_scale': 0.8, 'y_scale': 0.8})
            worksheet.set_row(i+1, 80)

            # Test
            # im = Image.open(BytesIO(requests.get(image_url).content))
            # st.write(im.size)

            data_bar.progress(((i+1)/n))

    #writer.save()
    writer.close()
    processed_data = output.getvalue()
    return processed_data

st.write("# Canvas Plus!")
st.write("This App supports extra features for Canvas.")

st.write('---')
token = st.text_input(label='Access token: ', key='token', value='')
url = 'https://canvas.nus.edu.sg/'
st.write('---')

if token != '':
    try:
        canvas = Canvas(url, token)
        courses = {course.name: course.id for course in canvas.get_courses()}
        course_name = st.selectbox('Courses:', courses, index=0)
        course = canvas.get_course(courses[course_name])

        tasks = ['Download Student List', 'Download Discussion Data']
        task = selected_task = st.selectbox('What to do:', tasks, index=0)

        st.write('---')

        if task == tasks[0]:

            # Filter with group categories
            is_cat_filter = st.checkbox('Filter with group categories', value=False)

            group_cats = {cat.name: cat for cat in course.get_group_categories()}
            if is_cat_filter:
                selected_cat = st.selectbox('Group Categories: ', group_cats)
            else:
                selected_cat = ''
            
            # Selected student information
            info = {'Avatar': 'avatar_url', 
                    'Student Number': 'integration_id', 
                    'Email': 'primary_email'}
            info_columns = st.multiselect('Student Info.: ', info, info)

            # Select group information
            default_cat = None if selected_cat == '' else selected_cat
            cat_columns = '+++'.join(st.multiselect('Group information: ', 
                                                    group_cats, default=default_cat,
                                                    key='cat_columns'))
            to_preview = st.button('Preview')
            
            if to_preview:

                df, profiles = gen_preview_table(token, course_name, 
                                                 selected_cat, info_columns, info,
                                                 cat_columns)


                df_xlsx = to_excel(token, course_name, selected_cat, info_columns, info, cat_columns)
                st.download_button('Download', df_xlsx, file_name= 'students.xlsx')

                st.write(df.to_html(escape=False), unsafe_allow_html=True)
        
        elif task == tasks[1]:

            thread = []
            name = []
            student_num = []
            date = []
            topics_dates = {}

            # only_pinned = st.checkbox('Only pinned topics', value=False)
            all_topics = course.get_discussion_topics()
            topic_dict = {t.__str__(): t for t in all_topics}
            selected_topics = st.multiselect('Topics', topic_dict, topic_dict)
            to_generate = st.button('Generate')

            if to_generate:
                for topic_name in selected_topics:
                    topic = topic_dict[topic_name]
                    topics_dates[topic.title] = topic.created_at
                    entries = list(topic.get_topic_entries())
                    n = len(entries)
                    st.write(f'{topic.__str__()}')
                    if n > 0:
                        data_bar = st.progress(0.0)
                    else:
                        data_bar = st.progress(1.0)
                    for i, a in enumerate(topic.get_topic_entries()):
                        try: 
                            sn = course.get_user(a.user_id).integration_id
                            thread.append(topic.title)
                            name.append(a.user_name)
                            student_num.append(sn)
                            date.append(a.updated_at_date)
                            for b in a.get_replies():
                                thread.append(topic.title)
                                name.append(b.user_name)
                                student_num.append(course.get_user(b.user_id).integration_id)
                                date.append(b.updated_at_date)
                        except:
                            st.write(f'{a.user_name} dropped the module.')
                        data_bar.progress(((i+1)/n))
            
                posts = pd.DataFrame({'Name': name, 'Number': student_num, 
                                      'Topics': thread, 'Date': date})

                st.write('#### Discussion Board Records: ')
                results = posts['Topics'].value_counts()
                results.name = 'Replies'
                st.write(results)

                posts['Date'] = posts['Date'].dt.tz_localize(None)
                # posts.to_excel('discussion.xlsx', index=False)

                output = BytesIO()
                writer = pd.ExcelWriter(output, engine='xlsxwriter')
                posts.to_excel(writer, index=False, sheet_name='Sheet1')
                # workbook = writer.book
                # worksheet = writer.sheets['Sheet1']
                # writer.save()
                writer.close()
                posts_xlsx = output.getvalue()

                st.download_button('Download', posts_xlsx, file_name= 'discussion.xlsx')

    except InvalidAccessToken:
        st.error('Invalid access token!')
