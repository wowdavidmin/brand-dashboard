# Heartbeat Tasks
## Jake 메일 확인
- `C:\Users\user\.openclaw\workspace\jake_mail_signal.txt` 파일이 존재하는지 확인
- 파일이 있으면 내용을 읽고 David에게 Telegram으로 아래 형식 알림:
  "📬 Jake로부터 새 메일 도착!
   발신자: [sender]
   제목: [subject]
   수신시간: [received]"
- 알림 보낸 후 시그널 파일 삭제
