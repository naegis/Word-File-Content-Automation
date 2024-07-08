import os
import openai
from docx import Document
from docx.shared import Pt
from docx.shared import RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# OPEN AI API
api_key = 'sk-proj-NGZNBp6g79fVGvNXH6HOT3BlbkFJ0SDMjIcE270uy6zAKyIZ'
openai.api_key = api_key

maincourse = "Interpersonal Communication Training Courses in Singapore"
country = "Singapore"
training_courses = [
    "Active Listening Skills Training Courses in Singapore"
]

# PROMPTING
def generate_content(prompt):
    try:
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=2000,
            n=1,
            stop=None,
            temperature=0.7,
        )
        return response.choices[0].message['content'].strip()
    except Exception as e:
        print(f"An error occurred: {e}")
        return None

# WORD MAKING
def paste_content():
    doc = Document()

    def heading(course):
        subtitle = doc.add_paragraph()
        subtitle_run = subtitle.add_run(course)
        subtitle_run.font.name = 'Arial'
        subtitle_run.font.size = Pt(20)
        subtitle_run.font.color.rgb = RGBColor(21,95,129)  # Light Blue for Subcourse Heading 2

    def line_splitter(prompt):
        # Split the prompt based on bullet points detection, in my case, i edited the prompt to add dashes at the start of the bullet points
        bullet_point = '- '  
        if bullet_point in prompt:
            parts = prompt.split(bullet_point, 1)
            description = parts[0].strip()
            bullet_list = parts[1].strip().split('\n')
        else:
            description = prompt.strip()
            bullet_list = []

        #adds an indentation 
        desc_paragraph = doc.add_paragraph()
        desc_paragraph.paragraph_format.left_indent = Pt(36)  # Adjust indent as needed
        desc_paragraph.add_run(description)

        for item in bullet_list:
            if item.strip():  # Check if bullet is not empty
                doc.add_paragraph(item.strip(), style='List Bullet')

    for course in training_courses:
        prompts = [
            f"Introduction Prompt: Please create a unique introduction for this title {course}. Lastly, please reiterate the title in the last sentence of the last paragraph. Make it 4-5 paragraphs.",
            f"Who Should Attend this {course} Prompt:Please create a 3 paragraph introduction For the training course title {course}. Lastly, please reiterate the title in the last sentence of the last paragraph. Can you also list down people who might be interested about this. (Use dashes before their names in place of bullets) Just up their titles/names no need for a description after the paragraphs. ",
            f"Course Duration Prompt: For the training course {course} title, create a 3 sentence paragraph introduction explaining the duration of the course training. Base the durations to the following: 3 full days, 1 day, half day, 90 minutes and 60 minutes. Please mention the title again in any of the 3 sentences.",
            f"Course Benefits Prompt: For the training course {course}, please create a 1 sentence description introducing the potential benefits of the training course. Afterwards, please list down 10 benefits of the course in bullet form. Use dashes before the benefits. ",
            f"Course Objectives Prompt Please create a 2 sentences description about the objectives of this course. Please mention the title again in any of the 2 sentences. Afterwards, in bullet form, (Use dashes before their names in place of bullets) please list down 12 objectives relating to the following benefits of the course. Make it different from the objectives. (insert the 10 course benefits)",
            f"Course Content Prompt Please create a 2 sentence description for the possible course content of {course}. Please put the course title somewhere in the sentences. Afterwards, create 12 sections using the same 12 objectives below, each with three topics composing of 1-2 sentences, pertaining to the following objectives of the course. Do not include colons, just the bullets. (insert the 12 objectives) Afterwards, create 12 sections (numbered list bold) using the same 12 objectives below, each with three topics (in bullet form) (Use dashes before their names in place of bullets) composing of 1-2 sentences, pertaining to the following objectives of the course. Do not include colons, just the bullets. >2 Full Days >9 a.m to 5 p.m (the bullets are constant, put it after the introduction)",
            f"Course Fees Prompt Please create a 3 sentence paragraph description for the possible course fees of (title). Please put the course title somewhere in the sentences. Specify that there will be 4 pricing options but do not specifically specify any. (insert these bullets after the introduction) USD 679.97 For a 60-minute Lunch Talk Session. USD 289.97 For a Half Day Course Per Participant. USD 439.97 For a 1 Day Course Per Participant. USD 589.97 For a 2 Day Course Per Participant. Discounts available for more than 2 participants. ",
            f"Upcoming and Brochure Download Prompt: Please create a 3 sentence paragraph description for the possible upcoming updates or to avail brochures about the training course {course}. Please put the course title somewhere in the sentences. "
        ]

        intro = generate_content(prompts[0])
        attendees = generate_content(prompts[1])
        duration = generate_content(prompts[2])
        benefits = generate_content(prompts[3])
        objectives = generate_content(prompts[4])
        content = generate_content(prompts[5])
        fees = generate_content(prompts[6])
        brochure = generate_content(prompts[7])

        if intro is None or attendees is None or duration is None or benefits is None or objectives is None or content is None or fees is None or brochure is None:
            print(f"Failed to generate content for {course}. Check error messages above.")
            continue

        title = doc.add_heading(f'{course}', level=1)
        title_format = title.runs[0].font
        title_format.size = Pt(26)

        #Add introductory paragraph with Arial and 1.5 line spacing
        intro_paragraph = doc.add_paragraph()
        intro_paragraph.paragraph_format.line_spacing = 1.5
        intro_run = intro_paragraph.add_run("Our training course \"Interpersonal Communication Training Courses in Singapore\" is also available in Orchard, Marina Bay, Bugis, Tanjong Pagar, Raffles Place, Sentosa, Jurong East, Tampines, Changi, and Woodlands.")
        intro_run.bold = True
        intro_run.font.name = 'Arial'

        doc.add_paragraph(intro)
        doc.paragraphs[-1].style = doc.styles['Normal']

        #Who Should Attend this [Sub course Title] Training Course in [Country] (Make this a subtitle but the format should be Arial, Font size: 20, Font color = blue)
        heading(f'Who Should Attend this {course}:')
        line_splitter(attendees)

        #Course Duration for [Sub course Title] Training Course in [Country] (Make this a subtitle but the format should be Arial, Font size: 20, Font color = blue)
        heading(f'Course Duration for {course}')
        duration_lines = duration.split('\n\n')
        intro_type = duration_lines[0]
        doc.add_paragraph(intro_type)
        doc.add_paragraph("2 Full Days", style='List Bullet')
        doc.add_paragraph("9 a.m. to 5 p.m.", style='List Bullet')

        #Course Benefits of [Sub course Title] Training Course in [Country]  
        heading(f'Course Benefits of {course}:')
        line_splitter(benefits)

        #Course Objectives of [Sub course Title] Training Course in [Country]  
        heading(f'Course Objectives of {course}:')
        objectives_lines = objectives.strip().split('\n\n')
        intro_objectives = objectives_lines[0]
        objectives_list = objectives_lines[1:]

        doc.add_paragraph(intro_objectives)
        doc.paragraphs[-1].style = doc.styles['Normal']

        for objective in objectives_list:
            if objective.startswith('1.'):
                p = doc.add_paragraph()
                p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                p.paragraph_format.line_spacing = 1.5
                p.style = 'List Number'
                p.add_run(objective).bold = True
                p.runs[0].font.name = 'Arial'
            else:
                doc.add_paragraph(objective, style='List Bullet')

        #Course Fees for [Sub course Title] Training Course in [Country]  
        heading(f'Course Fees for {course}:')
        fees_lines = fees.split('\n\n')
        fees_intro = fees_lines[0]
        doc.add_paragraph(fees_intro)
        doc.add_paragraph("USD 679.97 For a 60-minute Lunch Talk Session. ", style='List Bullet')
        doc.add_paragraph("USD 289.97 For a Half Day Course Per Participant. ", style='List Bullet')
        doc.add_paragraph("USD 439.97 For a 1 Day Course Per Participant. ", style='List Bullet')
        doc.add_paragraph("USD 589.97 For a 2 Day Course Per Participant. ", style='List Bullet')
        discount_paragraph = doc.add_paragraph()
        discount_run = discount_paragraph.add_run("Discounts available for more than 2 participants.")
        discount_run.bold = True
        discount_paragraph.style = 'List Bullet'

        #Upcoming Course and Course Brochure Download for [Sub course Title] Training Course in [Country]  
        heading(f"Upcoming Course and Course Brochure Download for {course}")
        brochure_lines = brochure.split('\n\n')
        brochure_intro = brochure_lines[0]
        #Add intro paragraph
        doc.add_paragraph(brochure_intro)


    
    # Save the document
    downloads_path = os.path.join(os.path.expanduser("~"), "Downloads", f"{course}.docx")
    try:
        doc.save(downloads_path)
        print(f"Document created successfully and saved to {downloads_path}")
    except Exception as e:
        print(f"An error occurred while saving the document: {e}")

# Execute the function to generate and save the document


paste_content()

