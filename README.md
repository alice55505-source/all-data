# all統計

一個可安裝的 PWA（漸進式網頁應用），用來彙整各召會上傳的週報，計算兒童與青職各項每週平均，並統計全台已上傳召會的總計。資料儲存在 Cloudflare D1（房間制），同一個房間代碼的人共用同一份資料，可跨裝置、跨團隊使用。

## 房間制（多人協作）

1. 開啟網站，點「建立新房間」會產生一個房間代碼（例如 `ABC123`）。
2. 把這個代碼告訴團隊，讓對方在首頁「輸入房間代碼加入」；房間列上的「複製房間代碼」可以直接複製代碼分享，不會複製成連結，避免轉傳訊息時不小心夾帶連結外流。
3. 同一個房間的人看到、上傳的都是同一份資料（召會分組、兒童／青職統計）；不同房間之間完全隔離，不知道代碼就進不去、看不到。
4. 房間列上可以按「✏️ 命名」幫房間取個好記的名字（例如「雲嘉區兒童統計」），同房間的人都會看到這個名字。
5. 房間內的資料每隔幾秒會自動跟伺服器同步一次，隊友剛上傳/修改的資料通常不用手動按「重新整理資料」就會自動出現；也可以隨時手動點「重新整理資料」立即拉取最新狀態。「離開房間」回到首頁。
6. 新建立的房間預設帶有雲東區、雲西區、嘉義區、民雄區、朴子區共 28 個召會的清單，可在設定裡自行修改。

資料存在你自己 Cloudflare 帳號底下的 D1 資料庫，不會經過任何第三方；房間本身沒有密碼保護，安全性完全取決於代碼本身不外流（不要把代碼貼在公開的地方）。

## 使用方式

上傳只有一個共用區塊：選召會、上傳一份「週報網格」格式的 `.xlsx`，系統會一次算出「兒童」「青職」兩個分頁的統計，不需要分開上傳兩次。

1. 選擇要上傳資料的召會。
2. 上傳該召會的週報 Excel 檔（`.xlsx`），需包含「週報網格」分頁（同時含學齡前、國小、青職等各年齡層欄位）。
3. 點擊「解析並加入」，系統會同時算出：
   - **兒童**：主日 =「兒童」小計欄 +「主日」類別的「國小」欄；召會生活、小排皆為該類別「學齡前」+「國小」的每週平均。
   - **青職**：主日、家聚會（出訪＋受訪）、小排、生命讀經皆為該類別「青職」欄位的每週平均。

兩個分頁（兒童／青職）各自顯示召會清單的上傳進度、依區域分組的結果表與全台總計，但共用同一次上傳的資料 —— 移除或清除某召會時，兩邊會一併移除。

### 設定召會清單（兩層：區域 → 召會）

點右上角「⚙ 設定召會」可自訂這個房間要追蹤的召會名單，分成「區域」「召會」兩層，兩個分頁共用同一份清單：

- 文字框裡，區域名稱獨立一行，該區底下的召會每行一個並加上縮排（例如兩個空格）
- 「清空全部」可以把名單清空，改成自己輸入的清單
- 「還原預設清單」可以隨時换回預設的 28 個召會（雲東區、雲西區、嘉義區、民雄區、朴子區共 5 個區域）
- 儲存後會即時套用給這個房間所有人

清單只是決定下拉選單、上傳進度、結果表與總計要涵蓋哪些召會；若把某個召會從清單移除，該召會先前上傳的資料不會被刪除，只是暫時不會顯示，把名稱加回清單就會再出現。

### 各召會結果表

每個分頁的「各召會結果」會依區域分組：每個區域最上面一列是該區域已上傳召會的**總計**（把區域內每個召會的數字加總，週數例外——週數顯示的是平均週數，僅供參考），底下再列出區域內每個召會自己的數字。表格可以按「下載 Excel」匯出成 `.xlsx` 檔（欄位跟畫面上顯示的一致，含區域總計列）。全台總計卡片維持加總全部已上傳召會的邏輯，不受區域分組影響。

## 安裝為 App（PWA）

以 HTTPS 部署後，桌面或手機瀏覽器可將此頁面「加入主畫面 / 安裝應用程式」，離線後仍可開啟介面（但需要連線才能讀寫房間資料）。

## 技術架構

- 前端：純靜態網頁（`index.html` / `css/style.css` / `js/app.js`），不需建置流程。
- 後端：Cloudflare Pages Functions（`functions/api/rooms/`），呼叫 Cloudflare D1 資料庫。
- 每個房間是 D1 `rooms` 表裡的一列：`name` 是房間名稱、`groups_json` 是召會分組、`stats_json` 是每個召會的兒童／青職統計，都是 JSON 字串。

上傳、移除召會、清除全部、改房間名稱，都是各自獨立的 API（`PUT/DELETE /api/rooms/:id/congregations/:name`、`POST /api/rooms/:id/reset`、`PATCH /api/rooms/:id/name`），伺服器收到請求時才讀取當下最新的資料做修改再寫回去，不是前端把整包資料整個蓋過去。這樣兩個人幾乎同時上傳不同召會時，才不會其中一人的存檔把另一人的蓋掉。只有「設定召會清單」是整批覆蓋（`PUT /api/rooms/:id/groups`），因為這是特意的、低頻率的整批編輯動作。

## 部署到 Cloudflare（首次設定）

這個專案需要一個 Cloudflare D1 資料庫，綁定到現有的 Cloudflare Pages 專案。以下指令請在本機（已 clone 這個 repo 的機器）執行：

```bash
# 1. 安裝相依套件（含 wrangler CLI）
npm install

# 2. 登入你的 Cloudflare 帳號（會開啟瀏覽器授權）
npx wrangler login

# 3. 建立正式的 D1 資料庫
npx wrangler d1 create all-children-db
```

第 3 步指令執行完會印出類似這樣的內容：

```
[[d1_databases]]
binding = "DB"
database_name = "all-children-db"
database_id = "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```

把印出的 `database_id` 貼到 repo 根目錄 `wrangler.toml` 裡取代 `REPLACE_WITH_DATABASE_ID`，然後：

```bash
# 4. 在正式的 D1 資料庫套用資料表結構
npx wrangler d1 execute all-children-db --remote --file=./schema.sql

# 5. 把 wrangler.toml 的改動 commit 起來（database_id 不是密鑰，可以放心 commit）
git add wrangler.toml
git commit -m "Configure D1 database id"
git push

# 6. 部署（如果 Cloudflare Pages 已經用 Git 整合自動部署，push 之後它會自動重新部署；
#    如果需要手動部署，執行：）
npx wrangler pages deploy .
```

部署完成後，打開網站，建立一個新房間測試看看是否能正常上傳、儲存、在另一台裝置/瀏覽器打開同一個房間連結確認資料同步。

### 更新既有的正式資料庫（migrations）

如果正式的 D1 資料庫是之前就建立過的（已經有房間資料在裡面），之後每次 `schema.sql` 以外新增的資料表欄位，會放在 `migrations/` 資料夾裡，需要手動套用一次：

```bash
npx wrangler d1 execute all-children-db --remote --file=./migrations/0002_add_room_name.sql
```

`--remote` 是操作正式環境的資料庫，請確認檔名是還沒套用過的再執行（同一個 migration 執行兩次通常不會出錯，`ALTER TABLE ... ADD COLUMN` 例外——欄位已存在時會報錯，屬正常現象，代表已經套用過了）。

### 本機開發（不動到正式資料庫）

```bash
npm install
npx wrangler d1 execute all-children-db --local --file=./schema.sql   # 建立本機測試用資料庫
npx wrangler pages dev .                                               # 啟動本機伺服器（預設 http://localhost:8788）
```

`--local` 只會操作本機的模擬資料庫，不會動到正式環境的資料。

## 目錄結構

```
index.html                                          主頁面（房間登入畫面 + 兒童／青職兩個分頁）
css/style.css                                        樣式（含深色模式）
js/app.js                                            前端邏輯：房間 API 串接、XLSX 解析、統計算繪、背景輪詢同步
js/xlsx.core.min.js                                  內嵌的 SheetJS 讀取 .xlsx 用（Apache-2.0，見 vendor/xlsx/LICENSE）
functions/api/_lib.js                                共用工具（不是路由，底線開頭的檔案 Pages Functions 不會當成 API 端點）
functions/api/rooms/index.js                         POST /api/rooms（建立房間）
functions/api/rooms/[id].js                          GET /api/rooms/:id（讀取房間完整資料）
functions/api/rooms/[id]/name.js                     PATCH /api/rooms/:id/name（設定房間名稱）
functions/api/rooms/[id]/groups.js                   PUT /api/rooms/:id/groups（整批覆蓋召會清單）
functions/api/rooms/[id]/reset.js                    POST /api/rooms/:id/reset（清除全部統計資料）
functions/api/rooms/[id]/congregations/[name].js     PUT/DELETE /api/rooms/:id/congregations/:name（單一召會的存檔／移除）
schema.sql                                           D1 資料表結構（新建資料庫用）
migrations/                                          既有資料庫要手動套用的欄位異動
wrangler.toml                                        Cloudflare Pages/D1 設定（含 database_id）
manifest.json                                        PWA manifest
sw.js                                                Service Worker（離線快取，不快取 /api/ 請求）
icons/                                                App 圖示
```
