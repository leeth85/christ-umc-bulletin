# Christ UMC Digital Bulletin

A free, serverless, mobile-friendly digital church bulletin web app customized in English for **Christ United Methodist Church** (`www.christumc.cc` · *Gather · Grow · Go*), adapted from the open-source MIT [Smart Jubo](https://github.com/swbts2006-crypto/smart-jubo) project.

## Key Features

| Tab | What Congregation Members Can Do |
|---|---|
| **Home** | Liturgical Christian Year banner (Advent, Christmas, Epiphany, Lent, Holy Week, Easter, Pentecost, Reign of Christ, Thanksgiving, Ordinary Time), **Livestream**, **Online Giving (Tithe.ly)**, **Connect Card**, **Prayer Request**, interactive **Order of Worship** (with bold congregational responses), top announcements, and worship schedule |
| **Sermon** | Sunday Scripture reading (`Matthew 6:6–10`), Sermon Notes, Reflection/Life Group questions, and a live font-size slider (`A-` / `A+`) |
| **Hymns** | Full **Hymns & Worship Lyrics** cards (`UMH #95` Doxology, opening/closing songs, optional CCLI license footer) **plus** optional sheet music / slide image uploads with pinch-to-zoom |
| **News** | Weekly church announcements with auto-clickable phone numbers, email addresses, and website links |
| **Connect** | **Digital Connect Card** (guest welcome & attendance by service: `8:00 AM`, `10:00 AM DR`, `10:00 AM YR`, `Online`), **Prayer Requests (Joys & Concerns)**, **Pastoral Care & Event RSVP**, **Seasonal Gratitude Journal**, and **Past Bulletins Archive** |
| **Admin** | Password-protected in-browser editor (`admin1234` initially) — edit everything and click **Download Deployment File (`index.html`)** |

## Folder Structure

```text
christ-umc-bulletin/
├── deploy/                        <-- Upload the files INSIDE this folder to Cloudflare Pages
│   ├── index.html                 Main app, UI, settings, and pre-loaded Nov 22, 2026 bulletin
│   ├── sw.js                      Offline reading & instant auto-update service worker
│   ├── recovery.html              Cache recovery utility
│   ├── manifest.json              Home-screen web app manifest
│   ├── _headers                   Cloudflare Pages cache revalidation rules
│   ├── icon-192.png               Cross & Flame home-screen icon (192x192)
│   └── icon-512.png               Cross & Flame home-screen icon (512x512)
├── GoogleSheets_Code.gs           Paste into Google Sheets > Extensions > Apps Script
├── Admin_and_Staff_Guide_EN.html  Full English manual for church staff (open & Print to PDF!)
└── LICENSE                        MIT License
```

## Quick Start (5 Minutes)

1. Create a free account at [cloudflare.com](https://www.cloudflare.com).
2. Go to **Workers & Pages → Create application → Pages → Upload assets**.
3. Name your project (e.g., `christumc-bulletin`) and drag **all 7 files inside `deploy/`** into the upload area, then click **Deploy site**.
4. Open your new URL, click **Admin** (top right), and enter `admin1234`.
5. Scroll down on **Home** to **[Setup] Church Profile & Links** to paste your direct **Tithe.ly Giving URL**, **Livestream URL**, and set a **new Admin Password**.
6. Click **Download Deployment File (index.html)** at the bottom of Home, replace `deploy/index.html` with the downloaded file, and upload the `deploy/` files via **Create new deployment**.

For full step-by-step instructions (including Google Sheets setup for Connect Cards & Prayer Requests), open **`Admin_and_Staff_Guide_EN.html`** in any browser (and click **Print / Save as PDF** if you want a PDF manual for staff).

---

## 한국어 요약 안내 (목사님 참고용)

* `deploy/` 폴더 안의 파일 7개를 **Cloudflare Pages**(무료)에 그대로 끌어다 놓으시면 즉시 `https://프로젝트이름.pages.dev` 주소가 생성됩니다. (교회 도메인 `bulletin.christumc.cc` 연결도 무료로 가능합니다.)
* 올려주신 **11월 22일 자 Christ UMC 주보**(설교 *"Teach Us How to Pray: Turn Your Mic On"*, 본문 *Matthew 6:6–10*, 예배 순서 `8:00 / 10:00 DR / 10:00 YR`, 교독문 볼드체 응답, 찬송 가사 탭, 6개 교회 공지사항, *Tithe.ly* 온라인 헌금 버튼, *Connect Card*, *Prayer Requests*, *Care & Event RSVP*)가 이미 초기 데이터로 모두 세팅되어 있습니다.
* 미국인 스태프나 행정 담당자분께는 `Admin_and_Staff_Guide_EN.html` 파일을 전달해 주시면 됩니다 (브라우저에서 열고 우측 상단 **Print / Save as PDF** 버튼을 누르면 깔끔한 영문 PDF 매뉴얼로도 바로 저장됩니다).
