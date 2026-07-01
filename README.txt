보소사진관 고객관리 최종 배포 파일입니다.

GitHub 새 저장소 또는 기존 boso-studio 연결 저장소에 아래 파일을 그대로 올리세요.
- index.html
- app.js
- styles.css
- price-list-1.png
- price-list-2.png

고객 기록까지 포함하려면 기존 사이트의 연동/백업 메뉴에서 JSON 백업을 받은 뒤,
파일 이름을 initial-data.json 으로 바꿔 이 폴더에 같이 올리세요.
주의: initial-data.json에는 고객명/전화번호/예약 내용이 들어가므로 공개 저장소에는 올리지 않는 것이 안전합니다.

Vercel 설정:
- Framework Preset: Other
- Root Directory: ./
- Build Command: 비움
- Output Directory: 비움
