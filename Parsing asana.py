import asana
import pandas as pd
import numpy as np
from datetime import datetime
startTime = datetime.now()

pd.set_option('display.max_rows', 7000)
pd.set_option('display.max_columns', 5000)
pd.set_option('display.width', 5000)

writer = pd.ExcelWriter('test_1.xlsx', engine='xlsxwriter')


personal_access_token = '1/1200243487618371:dd22fe988e2860a5ab98c95bc9a3c7d6'
client = asana.Client.access_token(personal_access_token)
me = client.users.me()
Workspace = me['workspaces'][0]
workspace_gid = Workspace['gid']
# ===========================================================================================================================================================

# Співробітники
members = client.users.get_users_for_workspace(Workspace['gid'], opt_fields='name')
data_members = pd.DataFrame(members)
data_members.to_excel(writer, sheet_name='Workers')
print("Час вигрузки співробітників ", datetime.now() - startTime)
# ===========================================================================================================================================================


# Проекти
projects = client.projects.get_projects_for_workspace(workspace_gid, opt_fields = ['created_at','name', 'due_date', 'due_on', 'archived', 'owner', 'completed'])
data_projects= pd.DataFrame(projects)

lst_owners = []
for i in range(len(data_projects)):
    try:
        lst_owners.append(data_projects['owner'][i].get('gid'))
    except:
        lst_owners.append(np.nan)
data_projects['owner'] = lst_owners
data_projects_merge = pd.merge(left=data_projects, right=data_members, left_on='owner', right_on='gid', how='left')
data_projects_merge.columns = ['Код проекту', 'Архівований', 'Статус', 'Дата створення', 'Виконати до', 'Срок', 'Назва', 'Код власника', 'g', 'Власник']
data_projects_merge.drop(['g'], axis=1, inplace=True)
data_projects_merge = data_projects_merge[data_projects_merge['Архівований'] == False]
data_projects_merge.reset_index(drop=True, inplace=True)
data_projects_merge.to_excel(writer, sheet_name='Projects')
print("Час вигрузки проектів ", datetime.now() - startTime)
# ===========================================================================================================================================================


# Задачі
data_task = pd.concat([pd.DataFrame(client.tasks.get_tasks_for_project(data_projects_merge['Код проекту'][i], {'followers':'gid'},
                                                                       opt_fields=['created_at','name', 'due_date', 'due_on','assignee',
                                                                        'completed', 'projects'])) for i in range(len(data_projects_merge))], ignore_index=True)

name_projects = data_projects_merge[['Код проекту', 'Назва']]
lst_gid_projects = []
lst_gid_assignee = []
for i in range(len(data_task)):
    for j in data_task['projects'][i]:
        try:
            lst_gid_projects.append(j.get('gid'))
        except:
            lst_gid_projects.append(np.nan)
for i in range(len(data_task)):
    try:
        lst_gid_assignee.append(data_task['assignee'][i].get('gid'))
    except:
        lst_gid_assignee.append(np.nan)

data_task['projects'] = lst_gid_projects
data_task['assignee'] = lst_gid_assignee
data_tasks_merge = pd.merge(left=data_task, right=name_projects, left_on='projects', right_on='Код проекту', how='left')
data_tasks_merge = pd.merge(left=data_tasks_merge, right=data_members, left_on='assignee', right_on='gid', how='left')
data_tasks_merge.drop(['projects', 'gid_y', 'assignee'], axis=1, inplace=True)
data_tasks_merge.columns = ['Код задачі' , 'Статус', 'Дата створення', 'Виконати до', 'Назва задачі', 'Код проекту', 'Назва проекту', 'Власник']
data_tasks_merge.to_excel(writer, sheet_name='Tasks')
print("Час вигрузки задач ", datetime.now() - startTime)
writer.save()
# ===========================================================================================================================================================


#Підзадачі
data_subtasks = pd.concat([pd.DataFrame(client.tasks.get_subtasks_for_task(data_task['gid'][i],
                                                                opt_fields=['created_at','name', 'due_date', 'due_on',
                                                                            'assignee', 'completed', 'tasks','parent'])) for i in range(len(data_tasks_merge))],ignore_index=True)
lst_gid_tasks = []
lst_gid_assignee_sbt = []
name_tasks = data_tasks_merge[['Код задачі', 'Назва задачі']]
for i in range(len(data_subtasks)):
    try:
        lst_gid_assignee_sbt.append(data_subtasks['assignee'][i].get('gid'))
    except:
        lst_gid_assignee_sbt.append(np.nan)
for i in range(len(data_subtasks)):
    try:
        lst_gid_tasks.append(data_subtasks['parent'][i].get('gid'))
    except:
        lst_gid_tasks.append(np.nan)
data_subtasks['assignee'] = lst_gid_assignee_sbt
data_subtasks['parent'] = lst_gid_tasks
data_subtasks_merge = pd.merge(left=data_subtasks, right=name_tasks, left_on='parent', right_on='Код задачі', how='left')
data_subtasks_merge = pd.merge(left=data_subtasks_merge, right=data_members, left_on='assignee', right_on='gid', how='left')
data_subtasks_merge.drop(['parent', 'gid_y', 'assignee'], axis=1, inplace=True)
data_subtasks_merge.columns = ['Код підзадачі', 'Статус', 'Дата створення', 'Виконати до', 'Назва підзадачі', 'Код задачі', 'Назва задачі', 'Власник']
data_subtasks_merge.to_excel(writer, sheet_name='sub_tasks')
print("Час вигрузки підзадач одніє задачі ", datetime.now() - startTime)




