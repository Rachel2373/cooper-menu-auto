from mailmerge import MailMerge

template = 'Menuauto.docx'
document = MailMerge(template)

# print(document.get_merge_fields())

#year = '2019'
#month = '7月'
#day = '22'
#meal_zh = '晚餐'
#lesson_zh = '2'
#lesson_en = "'Get Cooper23nt Experience Training Course"

protein_1_zh = '3'
protein_2_zh = '贝 - 清蒸123123'
protein_3 = '鱼 23鱼'

fruit_2 = '1'

tips_1_zh = '避免“短平快”的43会越减越肥！）'
tips_1 = '123212312546igjfn gets more weight'
tips_2_zh = '不吃或少吃加工后的（升糖快，再次饥饿快，还有隐形糖和添加剂）'
meal = 'dinner'
soup = 'Seaweed and Egg Soup'
fruit_2_zh = '香瓜'
tips_2 = 'Avoid processed food, which has higher glycemic index, plus with hidden sugar and additive'
fruit_3 = 'Pitaya'
fruit_1_zh = '哈112345密瓜'
fruit_1 = 'Hami Melon'
soup_zh = '紫菜蛋花汤'
protein_2 = 'Shellfish - Steamed Scallops'
protein_1 = 'Chicken - Curry  Chicken'
fruit_3_zh = '火龙果'
protein_3_zh = '鱼 - 清蒸多宝鱼'

month_dict = {'1': 'January', '2': 'February', '3': 'March', '4': 'April', '5': 'May', '6': 'June', '7': 'July',
              '8': 'August', '9': 'September', '10': 'October', '11': 'November', '12': 'December'}
day_dict = {'1': '1st', '2': '2nd', '3': '3rd', '4': '4th', '5': '5th', '6': '6th', '7': '7th', '8': '8th', '9': '9th',
            '10': '10th',
            '11': '11th', '12': '12th', '13': '13th', '14': '14th', '15': '15th', '16': '16th', '17': '17th',
            '18': '18th', '19': '19th',
            '20': '20th', '21': '21st', '22': '22nd', '23': '23rd', '24': '24th', '25': '25th', '26': '26th',
            '27': '27th', '28': '28th',
            '29': '29th', '30': '30th', '31': '31st'}
meal_dict = {'早餐': 'breakfast menu', '午餐': 'lunch menu', '晚餐': 'dinner menu'}
protein_dict = {'虾':'shrink'}
fruit_dict = {}
soup_dict = {}
detail_dict = {}
nutrition = {}
food_dict = {}
tips_dict = {}



year = input('哪年:')
month = input('哪月:')
day = input('哪日:')
meal_zh = input('哪一餐:')
lesson_zh = input('课程中文名:')
lesson_en = input('课程英文名:')
print(protein_dict.keys())
protein_1_zh = input('蛋白质1:')
protein_2_zh = input('蛋白质2:')
protein_3_zh = input('蛋白质3:')

document.merge(year = year,
               month = month,
               day = day,
               meal_zh = meal_zh,
               meal_en = meal_dict[meal_zh],
               month_en = month_dict[month],
               day_en = day_dict[day],
               lesson_zh = lesson_zh,
               lesson_en = lesson_en,
               protein_1_zh = protein_1_zh,
               protein_1= protein_dict[protein_1_zh],
               protein_2_zh= protein_2_zh,
               protein_2= protein_dict[protein_2_zh],
               protein_3_zh= protein_3_zh,
               protein_3= protein_dict[protein_3_zh],

               fruit_2 = fruit_2,
               tips_1_zh=tips_1_zh,
               tips_1=tips_1,
               fruit_1_zh=fruit_1_zh,
               tips_2_zh=tips_2_zh,
               meal=meal_dict[meal_zh],
               soup=soup,
               fruit_2_zh=fruit_2_zh,
               tips_2=tips_2,
               fruit_3=fruit_3,
               fruit_1=fruit_1,
               soup_zh=soup_zh,


               fruit_3_zh=fruit_3_zh,)
document.write('demo2.docx')
