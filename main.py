import csv
import datetime

from pyPdf import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import reportlab

from mailjet_rest import Client

import base64

import yaml


def touch_pdf(template_filename, x, y, name):
    # https://stackoverflow.com/questions/1180115/add-text-to-existing-pdf-using-python
    packet = StringIO.StringIO()

    # create a new PDF with Reportlab
    can = canvas.Canvas(packet, pagesize=letter)
    can.setFillColorRGB(0, 0, 255)

    # print(can.getAvailableFonts())
    reportlab.rl_config.TTFSearchPath.append('/System/Library/Fonts/Supplemental')
    pdfmetrics.registerFont(TTFont('Zapfino', 'Zapfino.ttf'))

    can.setFont('Zapfino', int(round(300 / len(name))))  # Helvetica, see "Font Book"
    can.drawCentredString(x, y, name)
    can.save()

    # move to the beginning of the StringIO buffer
    packet.seek(0)
    new_pdf = PdfFileReader(packet)

    # read your existing PDF
    existing_pdf = PdfFileReader(file(template_filename, "rb"))
    output = PdfFileWriter()

    # add the "watermark" (which is the new pdf) on the existing page
    page = existing_pdf.getPage(0)
    page.mergePage(new_pdf.getPage(0))
    output.addPage(page)

    # finally, write "output" to a real file
    my_file_name = "./Attachments/Cert_" + name.replace(' ', '') + ".pdf"
    output_stream = file(my_file_name, "wb")
    output.write(output_stream)
    output_stream.close()

    print("PDF touch completed!")
    return my_file_name


def post_mail(my_api_key, my_secret_key, custom_id, my_post_email,
              my_send_from, send_to, my_cc, my_subject,
              my_msg_text=None, my_msg_html=None, my_file_name=None):
    data = {
        'Messages': [
            {
                "From": {
                    "Email": my_send_from
                },
                "To": [
                    {
                        "Email": send_to
                    }
                ],
                "Subject": my_subject,
                "CustomID": custom_id,
            }
        ]
    }

    if my_cc is not None:
        data['Messages'][0]['Cc'] = [
            {
                "Email": my_cc
            }
        ]

    if my_msg_text is not None:
        data['Messages'][0]['TextPart'] = my_msg_text

    if my_msg_html is not None:
        data['Messages'][0]['HtmlPart'] = my_msg_html

    if my_file_name is not None:
        with open(my_file_name, 'rb') as fil:
            b64_encoded = base64.b64encode(fil.read())

        data['Messages'][0]['Attachments'] = [
            {
                "ContentType": "application/pdf",
                "Filename": my_file_name,
                "Base64Content": b64_encoded
            }
        ]

    print('>> [{0}] - sending to `{1}`'.format(datetime.datetime.now(), send_to))
    if my_post_email:
        mailjet = Client(auth=(my_api_key, my_secret_key), version='v3.1')
        result = mailjet.send.create(data=data)
        print(result.status_code)
        print(result.json())


def get_students(my_file_name):
    students = {}
    i = 0
    with open(r'Recipients/' + my_file_name, 'rb') as f:
        reader = csv.reader(f, delimiter=',')

        for row in reader:
            if i != 0:  # skip the header line
                team_id, team_name, student_id, first_name, last_name, email, ind_code = \
                    row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip(), \
                    row[4].strip(), row[5].strip(), row[6].strip()

                # a student may use his/her guardian's email for registration, so email can duplicate!
                if first_name and last_name:
                    name = first_name + ' ' + last_name
                else:
                    print('>> [line {0}] name \'{1}\' or \'{2}\' is empty!'.format(i + 1, first_name, last_name))

                my_key = name.upper() + '~' + email.lower()
                if my_key in students.keys():
                    print('!! [line {0}] key "name~email" \'{1}\' duplicates!'.format(i + 1, my_key))
                    print('   Current line [' + str(i + 1) + ']: ', row)
                    print('   Existing individual: ', my_key, students[my_key])
                    print('')
                else:
                    students[my_key] = (name, email, ind_code, team_id, team_name, student_id,
                                        'line {0}'.format(i + 1))

            i += 1

    print('>> processed ' + str(i) + ' lines, including the first header line.')
    print('>> got ' + str(len(students.keys())) + ' students with non-empty distinct key of NAME~email.')
    assert len(students.keys()) == i - 1, '>> NOT ALL individual is counted!! expected: ' + str(i - 1) \
                                          + ' vs. actual: ' + str(len(students.keys())) + '.'
    return students


def get_teams():
    teams = {}
    i = 0
    with open(r'Recipients/Team List.csv', 'rb') as f:
        reader = csv.reader(f, delimiter=',')

        for row in reader:
            if i != 0:  # skip the header line
                team_name, team_id, team_code, first_name_1, last_name_1, first_name_2, last_name_2, \
                    first_name_3, last_name_3, first_name_4, last_name_4, email_1, email_2, email_3, email_4 = \
                    row[0].strip(), row[1].strip(), row[2].strip(), row[4].strip(), row[5].strip(), \
                    row[7].strip(), row[8].strip(), row[10].strip(), row[11].strip(), row[13].strip(), \
                    row[14].strip(), row[15].strip(), row[16].strip(), row[17].strip(), row[18].strip()

                if team_name:
                    teams[team_name] = (team_id, team_code, [(first_name_1, last_name_1, email_1),
                                                             (first_name_2, last_name_2, email_2),
                                                             (first_name_3, last_name_3, email_3),
                                                             (first_name_4, last_name_4, email_4)])

            i += 1

    print('>> processed ' + str(i) + ' lines, including the first header line.')
    print('>> got ' + str(len(teams.keys())) + ' teams with non-empty team-name.')
    return teams


def get_sent_student_ids():
    i = 0
    sent_student_ids = []
    with open(r'Logs/Mailjet_Email_Report_20210825225500.csv', 'rb') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            if i != 0:  # skip the header line
                if row[0] != 'cowconutsmath@gmail.com':
                    sent_student_ids.append(row[13][-4:])

            i += 1

    # print(sent_student_ids)
    # print(len(sent_student_ids))
    return sent_student_ids


def replace_tag(template, student, team):
    my_msg = template.replace('{name}', student[0].title())
    my_msg = my_msg.replace('{ind-code}', student[2])
    my_msg = my_msg.replace('{team-name}', student[4])

    my_msg = my_msg.replace('{team-id}', team[0])
    my_msg = my_msg.replace('{team-code}', team[1])

    for k in range(4):
        member = team[2][k]
        if member[0] and member[1] and member[2]:
            my_msg = my_msg.replace('{name' + str(1 + k) + '}', (member[0] + ' ' + member[1]).title())
            my_msg = my_msg.replace('{email' + str(1 + k) + '}', '(' + member[2] + ')<br>')
        else:
            my_msg = my_msg.replace('{name' + str(1 + k) + '}', '')
            my_msg = my_msg.replace('{email' + str(1 + k) + '}', '')

    return my_msg


def email_confirmation(post_email):
    send_from = 'cowconutsmath@gmail.com'
    cc = 'cowconutsmath@gmail.com'
    subject = '2021 Annual ShengMeng Math Competition - Confirmation'

    my_students = get_students('Individual List_8 Teams.csv')
    my_teams = get_teams()
    my_sent_student_ids = get_sent_student_ids()

    with open(r'./Templates/ConfirmationEmail.html', 'r') as f:
        message_html = f.read()

    j, k, start, end = 0, 0, 0, 330
    for my_key, value in my_students.items():  # 316 items
        if j < start or value[5] in my_sent_student_ids:
            j += 1
            continue
        if j >= end:
            break
        print('[info] j = ' + str(j))
        # print(value)

        msg = replace_tag(message_html, value, my_teams[value[4]])

        # no need to do 'Cc' because mailjet has email report
        post_mail(api_key, secret_key, 'confirmation-005', post_email,
                  send_from, value[1], None, subject + ' - ' + value[5],
                  None, msg, None)
        j += 1
        # if j > 1:       # for testing the email-sending functionality
        #     break
        k += 1

    # print('>> k = {0}, so will post {0} messages.'.format(k, k))


def email_participation(my_post_email):
    send_from = 'cowconutsmath@gmail.com'
    subject = 'Participation Certificate 2021'
    message_text = 'Thank you for attending the 4th Annual ShengMeng Math on Sept. 5th, 2021!'

    with open(r'Recipients/NameList.csv', 'rb') as f:
        reader = csv.reader(f, delimiter=',')
        for row in reader:
            print('for', row[0], ' sending to', row[1])
            file_name = touch_pdf('Templates/Participation Certificate_2021.pdf', 300, 560, row[0])

            post_mail(api_key, secret_key, 'test-participation-001', my_post_email,
                      send_from, row[1], None, subject, message_text, None, file_name)


def validate_0822():
    students = {}
    i = 0
    with open(r'Recipients/0822 individual.csv', 'rb') as f:
        reader = csv.reader(f, delimiter=',')

        for row in reader:
            if i != 0:  # skip the header line
                individual_id, first_name, last_name, email, team_name, team_id = \
                    row[0].strip(), row[1].strip(), row[2].strip(), row[4].strip(), row[5].strip(), row[6].strip()

                # a student may use his/her guardian's email for registration, so email can duplicate!
                if first_name and last_name:
                    name = first_name + ' ' + last_name
                else:
                    print('>> [line {0}] name \'{1}\' or \'{2}\' is empty!'.format(i + 1, first_name, last_name))

                my_key = name.upper() + '~' + email.lower()
                if my_key in students.keys():
                    print('!! [line {0}] key "name~email" \'{1}\' duplicates!'.format(i + 1, my_key))
                    print('   Current line [' + str(i + 1) + ']: ', row)
                    print('   Existing individual: ', my_key, students[my_key])
                    print('')
                else:
                    students[my_key] = (individual_id, name, email, team_name, team_id, 'line {0}'.format(i + 1))

            i += 1

        print('>> processed ' + str(i) + ' lines, including the first header line.')
        print('>> got ' + str(len(students.keys())) + ' students with non-empty distinct key of NAME~email.')
        assert len(students.keys()) == i - 1, '>> NOT ALL individual is counted!! expected: ' + str(i - 1) \
                                              + ' vs. actual: ' + str(len(students.keys())) + '.'


if __name__ == '__main__':
    with open(r'application.yml', 'r') as f:
        config = yaml.safe_load(f)

    api_key = config['mailjet'][0]['api_key']
    secret_key = config['mailjet'][0]['secret_key']

    # validate_0822()

    # get_sent_addresses()

    email_confirmation(post_email=config['debug']['post_email'])

    # email_participation(my_post_email=config['debug']['post_email'])
