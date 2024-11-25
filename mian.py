from pptx import Presentation
from pptx.util import Inches
from zhipuai import ZhipuAI
from pptx.dml.color import RGBColor

client = ZhipuAI(api_key="622b7be459b291e1fca40f128ae2447c.2tCWCQFAvP5a5ma2")  # 请填写您自己的APIKey


keyword = "ai,应用"
input_content = "ai创新应用"

# 拼接 input 到 title_query
title_query = f"用户想根据{input_content}写一份ppt，你先为他写一个明确的标题，10个字以内，不要给出多个选择，就是一个标题。"

title_response = client.chat.completions.create(
    model="glm-4-flash",  # 请填写您要调用的模型名称
    messages=[
        {"role": "user", "content": title_query},
    ],
)
title = title_response.choices[0].message.content
print(f"title:{title}")
# print(response.choices[0].message.content)

content_query = f'''这是用户的输入：{input_content}
他要就其中的主题写PPT汇报，在PPT之前需要有一份清晰的思路，请你帮他写一份清晰的创作思路，要求尽可能详尽，参考示例的输出。

示例：
输入：我要做一个培训项目，主题是《安全用电知识》受众是中小学生
输出：
1. 理解受众：中小学生应该掌握基础的安全用电常识，内容需要贴近学生的理解水平，避免过于抽象的技术术语。
2. 安全原则：介绍安全用电的基本规则，如何预防触电事故。
3. 实际操作：结合实际情况，讲述在家中和学校的安全用电行为。
4. 案例教育：引入一些典型的案例，让学生通过事例了解不安全用电的后果。
5. 常见误区：澄清一些常见的用电误区，提升学生的自我保护意识。
6. 互动环节：设置问答或小组讨论环节，强化学生对于知识点的掌握。'''

content_response = client.chat.completions.create(
    model="glm-4-flash",  # 请填写您要调用的模型名称
    messages=[
        {"role": "user", "content": content_query},
    ],
)

content = content_response.choices[0].message.content
print(f"content:{content}")

md_query = f'''根据主题：{keyword} 和标题{title}，参考思路{content}为用户制作一份PPT大纲，以markdown格式输出。大纲必须包括类似示例中对于小标题的内容扩充，不要求详细但需要针对小标题内容简述。

markdown格式大纲示例：

# 安全用电知识讲座

## 开场
- 引入主题：通过案例说明用电安全的重要性。
- 介绍重要性和必要性：阐述正确用电对个人和家庭的意义以及避免事故的必要性。

## 第一章：电的基本认识
- 电的用途：简述电在日常生活中的应用，如照明、加热、通讯等方面。
- 电的潜在危险：讲解不当使用电可能引起的风险，例如触电、火灾等。

## 第二章：安全用电原则
- 不私拉乱接电线：说明私拉乱接电线的危害和如何避免。
- 不在潮湿环境中使用电器设备：讲解潮湿对电器设备的影响及安全使用的方法。
- 正确使用电器设备：列举正确操作电器的要点。
- 不玩弄插座和开关：强调插座和开关使用的注意事项。

## 第三章：认识标识与注意事项
- 认识安全用电标识：解释安全标识的含义及其重要性。
- 注意远离高压设备：强调高压电的危险性和保持安全距离。
- 发现问题如何应对：讨论遇到电气问题的正确响应方式。

## 第四章：家庭中的安全用电
- 使用电器的正确方法：介绍家用电器使用的安全指南。
- 如何防止电器过载发热：探讨电器过载的原因及预防措施。
- 学会使用保险箱、漏电保护器：讲解如何使用这些设备预防电气事故。

## 第五章：校园中的安全用电
- 学校电线及设备的安全使用：指出学校用电设备的安全要点。
- 遇到雷雨天气如何安全用电：提供雷雨天气下的用电安全建议。
- 电教室、图书馆用电安全：讲述公共场所用电安全管理。

## 第六章：触电急救措施
- 认识触电现象：说明触电是如何发生的。
- 家庭中触电急救步骤：提供家庭触电事故的应急处理步骤。
- 学校中触电急救步骤：指导在学校发生触电事故的急救措施。

## 第七章：我应该如何做
- 学生在家庭中如何安全用电：给出学生在家中用电的指导建议。
- 学生在校园中如何安全用电：讨论学生在校园的安全用电行为。
- 安全用电的小提示：总结一些简单易行的用电安全小常识。

## 结束语
- 重要性的再次强调：总结讲座内容，重申用电安全的重要性。
- 鼓励学生安全的用电习惯：激励学生养成良好的用电习惯，确保自身安全。

## 互动环节
- 提问与答疑：解答学生在讲座中提出的问题。
- 分组讨论：我们可以为安全用电做些什么？引导学生思考和讨论在日常生活中如何实践安全用电。'''

# 创建聊天完成的响应，使用指定的模型对用户消息进行回复
md_response = client.chat.completions.create(
    model="glm-4-flash",  # 请填写您要调用的模型名称
    messages=[
        {"role": "user", "content": md_query},
    ],
)

md = md_response.choices[0].message.content
print(f"md:{md}")

# 解析Markdown内容，提取标题和对应的内容
sections = []
current_section = {'title': None, 'content': []}

lines = md.split('\n')
for line in lines:
    if line.startswith('## '):  # 二级标题
        if current_section['title'] is not None:
            sections.append(current_section)
        current_section = {'title': line[3:].strip(), 'content': []}
    elif current_section['title'] is not None:
        current_section['content'].append(line.strip())

if current_section['title'] is not None:
    sections.append(current_section)

# 创建一个新的PPT对象
ppt = Presentation()

# 遍历每个部分，创建新的幻灯片
for section in sections:
    # 添加一个新的幻灯片，并选择索引为5的幻灯片布局
    slide = ppt.slides.add_slide(ppt.slide_layouts[6])

    # 删除所有占位符
    for shape in slide.shapes:
        if shape.has_text_frame:
            shape.text_frame.clear()

    # 定义文本框的位置和大小
    left = top = Inches(1)
    width = height = Inches(5)

    # 在幻灯片上添加文本框
    txBox = slide.shapes.add_textbox(left, top, width, height)

    # 获取文本框的文本框对象
    tf = txBox.text_frame

    # 设置文本框的文本内容
    tf.text = section['title']

    # 添加内容
    for content in section['content']:
        p = tf.add_paragraph()
        p.text = content

# 保存PPT文件
ppt.save('test_ppt.pptx')