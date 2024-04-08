import yagmail
import pandas as pd
from datetime import datetime

start = datetime.now() #To calculate the time

sender = "your_mail@gmail.com"
yag = yagmail.SMTP(user= sender, password= "your_password_here") #Get the Password from your Google account -> App password

df = pd.read_excel("data.xlsx", "RAW")
# print(df) -> Make sure to check your data. Clean/manipulate if needed.

for index, row in df.iterrows():
    receiver = row['Email']
    name = row['Name']
    cid = row ['CID']
    uid = row ['UID']
    session = row ['Session']
    certificate = row ['Certificate']
    issued = row ['Issue Date'].date() #Taking only the date.
    pdf = row ['File']

    subject = f"Congratulations! {name}, Here is your {certificate} certificate"

    contents = [f"""
Hi there, {name}, Session {session}, GTCAC ID {cid}

Congratulations! You have successfully completed the {certificate} as of {issued}. Your certificate is attached to this email. Please download it and we would appreciate it if you could share your experience with us.

We look forward to your continued involvement with the club’s operations.

Certificate ID: {uid}
Verify Certificate: https://5jp20q-my.sharepoint.com/:x:/g/personal/gtcac_5jp20q_onmicrosoft_com/Efiz6QE9dz9OgcVYHSPlAXgBFx2JW4xLFtItzE5maKO3Lg?e=ZA88tZ

Sincerely,
Shaikh Ahmed
LinkedIn: https://www.linkedin.com/in/shaikhahmedbd
""", pdf]
        
    yag.send(to= receiver, subject=subject, contents=contents)    
    print(f"Email sent! to {name} with {pdf}")

print()
print("ALL EMAIL SENT!")

end = datetime.now()

print()
diff = end - start 
print(f"Total time took: {diff}") #To check how much time it took.