# SourceLab Toolbox

Free all-in-one online tools (PDF, images, video, audio, data, developer tools, calculators, link shortener, 24-hour file hosting). No login.

Built by Abhigyan. Telegram: https://t.me/SourceLabPro

## Run locally

```
docker build -t toolbox .
docker run -p 10000:10000 toolbox
```

Optional environment variables: `LINKS_SB_URL`, `LINKS_SB_KEY`, `LINKS_GATE` (persistent short links), `MAX_MB` (upload limit, default 20).
