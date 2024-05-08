import requests
from bs4 import BeautifulSoup
import pandas as pd
import re

def read_secrets_file(file_path):
    secrets = {}
    with open(file_path, 'r') as f:
        for line in f:
            # 주석 또는 빈 줄은 건너뜁니다.
            if line.strip().startswith('#') or not line.strip():
                continue
            # '='를 기준으로 키와 값을 나눕니다.
            key, value = line.strip().split('=', 1)
            # 키와 값을 딕셔너리에 추가합니다.
            secrets[key.strip()] = value.strip().strip('" ')
    return secrets

# 파일 경로
file_path = 'secret.txt'

# 시크릿 데이터 읽기
secrets = read_secrets_file(file_path)

# 액세스 토큰을 요청할 URL
token_url = "https://login.microsoftonline.com/1b071189-0ecb-4f7a-b453-9457c489fdde/oauth2/v2.0/token"

# 헤더 설정
token_headers = {
    "Content-Type": "application/x-www-form-urlencoded"
}

# 요청 데이터 설정
token_data = {
    "client_id": secrets['client_id'],
    "scope": "https://graph.microsoft.com/.default",
    "client_secret": secrets['client_secret'],
    "grant_type": "client_credentials"
}

# 액세스 토큰 요청
response = requests.post(token_url, headers=token_headers, data=token_data)

# 응답에서 액세스 토큰 추출
access_token = response.json()['access_token']

# Outlook 메일함의 메시지를 읽기 위한 URL
outlook_mail_url = f"https://graph.microsoft.com/v1.0/users/{secrets['user_name']}/mailFolders/inbox/messages?$top=50"

# 헤더 설정
outlook_headers = {
    "Authorization": "Bearer " + access_token
}

# Outlook 메일함의 메시지를 가져오기 위한 GET 요청
outlook_response = requests.get(outlook_mail_url, headers=outlook_headers)

# 결과를 JSON 형식으로 파싱
mail_data = outlook_response.json()

# 정규 표현식을 사용하여 호스트네임 추출
def extract_vm_hostname(content):
    match = re.search(r"/virtualMachines/(\w+)", content)
    if match:
        return match.group(1)
    else:
        return None
    
# "Out of Memory"와 "VM Resource Health"를 포함하는 메일만 추출하여 리스트에 추가
target_mails = []
for message in mail_data['value']:
    subject = message['subject']
    receivedDateTime = message['receivedDateTime']
    content = BeautifulSoup(message['body']['content'], 'html.parser').get_text()
    if "Out of Memory" in subject or "VM Resource Health" in subject or "Out of Memory" in content or "VM Resource Health" in content:
        vm_hostname = extract_vm_hostname(content)
        target_mails.append({
            "receivedDateTime": receivedDateTime,
            "subject": subject,
            "content": content,
            "VM" : vm_hostname
        })
    
# 기존 엑셀 파일 경로
existing_excel_file = "[HCM-305] Hybrid_Cloud_MTP - CSP 서비스 Monitoring 방안수립 - 동보관리.xlsx"

# 추출된 메일을 데이터프레임 형식으로 저장
mail_df = pd.DataFrame(target_mails)

# 칼럼 순서 설정
mail_df = mail_df[['receivedDateTime', 'VM', 'subject', 'content']]

# 기존 엑셀 파일 열기
with pd.ExcelWriter(existing_excel_file, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
    # 메일 내용에 따라서 시트 선택
    if any("Azure" in content for content in mail_df["content"]):
        target_sheet_name = "Azure_temp"
    elif any("OCI" in content for content in mail_df["content"]):
        target_sheet_name = "OCI"
    elif any("AWS" in content for content in mail_df["content"]):
        target_sheet_name = "AWS"
    
    # 기존 시트 불러오기
    try:
        existing_df = pd.read_excel(existing_excel_file, sheet_name=target_sheet_name)
    except FileNotFoundError:  # 파일이 없으면 빈 데이터프레임 생성
        existing_df = pd.DataFrame()
    
    # 새로운 데이터프레임을 기존 데이터프레임에 추가
    merged_df = pd.concat([existing_df, mail_df], ignore_index=True)
    
    # 데이터프레임을 지정된 시트에 쓰기
    merged_df.to_excel(writer, index=False, sheet_name=target_sheet_name)

# # 추출된 메일 출력
# for mail in target_mails:
#     print("제목:", mail['subject'])
#     print("내용:", mail['content'])
#     print("받은 시간:", mail['receivedDateTime'])
#     print("---------------------------")