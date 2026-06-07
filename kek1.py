from imapclient import IMAPClient
from email import message_from_bytes

HOST = 'imap.gmail.com'
USER = ''
PASSWORD = ''
list_for_stuff = []

with IMAPClient(HOST, ssl=True) as client:
    client.login(USER, PASSWORD)
    print("Успешное подключение!")
    client.select_folder("INBOX")
    messages = client.search(
        ["SUBJECT", "лента",
         "SINCE", "01-may-2026"],
        charset="UTF-8"
    )
    uid = messages[-1]
    data = client.fetch([uid], ["RFC822"])
    raw = data[uid][b"RFC822"]
    # print(raw.decode("utf-8", errors="replace"))  это не трогаем, это простой вариант
    msg = message_from_bytes(raw)

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()

            if ctype == "text/plain":
                text = part.get_payload(decode=True)

                charset = part.get_content_charset() or "utf-8"

                print(text.decode(charset, errors="replace"))
                list_for_stuff += [text.decode(charset, errors="replace")]

    else:
        text = msg.get_payload(decode=True)

        charset = msg.get_content_charset() or "utf-8"

        print(text.decode(charset, errors="replace"))
        list_for_stuff += [text.decode(charset, errors="replace")]

# with open('kekkek.txt', 'w', encoding='utf-8') as file:
#     for line in file:
print(list_for_stuff)