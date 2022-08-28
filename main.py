from mailbot import mailbot
import excel
import os

answer = ''
filepath = os.path.join(os.getcwd(), 'test_excel.xlsx')

while answer != '0':
    answer = input(
        '''원하는 작업을 선택하세요. \n[1] 테스트 이메일 발송 \n[2] 엑셀파일 수정 \n[0] 종료\n: ''')    
    print(f'{answer}번 작업을 선택하셨습니다.')
    if answer == '1':
        # 메일 보내기 테스트
        bot = mailbot()
        subject = '테스트이구요'
        contents = '보내지기만 하면 됩니다.'
        attach = os.path.join(os.getcwd(),'첨부파일테스트.txt')
        bot.send_mail(bot.bot, bot.info, subject, contents, 'whoami121212@gmail.com', path = attach)
        print('테스트 이메일 발송 완료')
    elif answer == '2':
        excel.FixingFile(filepath, 30)
        print('엑셀파일 수정작업이 완료되었습니다.')


