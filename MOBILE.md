# 手機使用方式

這個網頁版已經支援手機瀏覽器介面、前後鏡頭切換、PWA manifest、相機姿態辨識與 Excel 匯出。

## 本機測試

電腦上可以用：

```bash
./mp39_env/bin/python -m http.server 8001 --directory web
```

然後在同一台電腦開：

```text
http://127.0.0.1:8001
```

## 手機測試重點

手機不能用 `127.0.0.1` 連到你的電腦，因為手機上的 `127.0.0.1` 指的是手機自己。

如果要用手機連電腦，需要讓手機和電腦在同一個 Wi-Fi，並用電腦的區網 IP：

```text
http://你的電腦IP:8001
```

但是多數手機瀏覽器只允許在 HTTPS 或 localhost 使用相機。因此用區網 `http://` 開啟時，可能會無法啟動相機。

## 推薦正式方式

最穩定的方式是部署成 HTTPS 靜態網站，例如：

- GitHub Pages
- Netlify
- Vercel
- Cloudflare Pages

部署後用手機開 HTTPS 網址，瀏覽器就能正常要求相機權限。

## Excel 儲存

桌面 Chrome/Edge 若支援資料夾選擇，會跳出資料夾選擇。

手機 Safari/Chrome 通常不支援資料夾選擇，會改成直接下載 Excel 檔案。
