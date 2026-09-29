# 测试视频文件

由 `generate_test_videos.ps1` 使用工作区自带 `tools/ffmpeg.exe` 合成，全部为 **10 秒**、**30fps**、画面为 `testsrc2` 彩条+计时器，音轨为 **440Hz 正弦音（立体声）**。
片源完全由 lavfi 合成，不依赖任何外部素材。共 15 种容器/编码组合 × 2 档分辨率 = **30 个文件**，总体积约 **217 MB**。

## 生成 / 复跑

```powershell
# 全量生成（覆盖同名文件）
powershell -ExecutionPolicy Bypass -NoProfile -File .\generate_test_videos.ps1

# 只生成某几种格式 / 某一档分辨率
powershell -ExecutionPolicy Bypass -NoProfile -File .\generate_test_videos.ps1 -Formats mp4_h264,webm_vp9 -Resolutions 720p

# 只校验已有文件（不重新编码）
powershell -ExecutionPolicy Bypass -NoProfile -File .\generate_test_videos.ps1 -VerifyOnly

# 自定义参数
powershell -ExecutionPolicy Bypass -NoProfile -File .\generate_test_videos.ps1 -Duration 30 -Fps 25 -OutDir D:\tmp\videos
```

脚本参数：`-Ffmpeg`（默认 `..\tools\ffmpeg.exe`）、`-OutDir`、`-Duration`、`-Fps`、`-Resolutions`、`-Formats`（支持通配符）、`-SkipVerify`、`-VerifyOnly`。

## 文件清单

| 文件名 | 容器 / 视频编码 / 音频编码 | 分辨率 | 时长 | 体积 |
|---|---|---|---|---|
| test_720p_mp4_h264.mp4 | MP4 / H.264 (libx264 CRF23) / AAC 128k | 1280x720 | 10.00s | 3.88 MB |
| test_1080p_mp4_h264.mp4 | MP4 / H.264 / AAC 128k | 1920x1080 | 10.00s | 7.75 MB |
| test_720p_mp4_hevc.mp4 | MP4 / H.265 HEVC (libx265 CRF28, hvc1) / AAC 128k | 1280x720 | 10.00s | 2.61 MB |
| test_1080p_mp4_hevc.mp4 | MP4 / H.265 HEVC / AAC 128k | 1920x1080 | 10.00s | 5.86 MB |
| test_720p_mkv_h264.mkv | Matroska / H.264 / AAC 128k | 1280x720 | 10.02s | 3.88 MB |
| test_1080p_mkv_h264.mkv | Matroska / H.264 / AAC 128k | 1920x1080 | 10.02s | 7.75 MB |
| test_720p_mkv_vp9.mkv | Matroska / VP9 (CRF32) / Opus 96k | 1280x720 | 10.01s | 4.21 MB |
| test_1080p_mkv_vp9.mkv | Matroska / VP9 / Opus 96k | 1920x1080 | 10.01s | 8.77 MB |
| test_720p_webm_vp9.webm | WebM / VP9 (CRF32) / Opus 96k | 1280x720 | 10.01s | 4.21 MB |
| test_1080p_webm_vp9.webm | WebM / VP9 / Opus 96k | 1920x1080 | 10.01s | 8.77 MB |
| test_720p_webm_vp8.webm | WebM / VP8 (CRF10) / Vorbis q4 | 1280x720 | 10.00s | 1.05 MB |
| test_1080p_webm_vp8.webm | WebM / VP8 / Vorbis q4 | 1920x1080 | 10.00s | 1.90 MB |
| test_720p_avi_mpeg4.avi | AVI / MPEG-4 (q4) / MP3 128k | 1280x720 | 10.03s | 9.81 MB |
| test_1080p_avi_mpeg4.avi | AVI / MPEG-4 / MP3 128k | 1920x1080 | 10.03s | 21.15 MB |
| test_720p_mov_h264.mov | QuickTime / H.264 / AAC 128k | 1280x720 | 10.00s | 3.88 MB |
| test_1080p_mov_h264.mov | QuickTime / H.264 / AAC 128k | 1920x1080 | 10.00s | 7.75 MB |
| test_720p_mov_prores.mov | QuickTime / Apple ProRes **Proxy** / PCM s16le | 1280x720 | 10.00s | 18.69 MB |
| test_1080p_mov_prores.mov | QuickTime / Apple ProRes Proxy / PCM s16le | 1920x1080 | 10.00s | 28.13 MB |
| test_720p_ts_h264.ts | MPEG-TS / H.264 / AAC 128k | 1280x720 | 10.02s | 4.04 MB |
| test_1080p_ts_h264.ts | MPEG-TS / H.264 / AAC 128k | 1920x1080 | 10.02s | 7.99 MB |
| test_720p_flv_h264.flv | FLV / H.264 / MP3 128k | 1280x720 | 10.07s | 3.48 MB |
| test_1080p_flv_h264.flv | FLV / H.264 / MP3 128k | 1920x1080 | 10.07s | 6.94 MB |
| test_720p_asf_wmv2.wmv | ASF / WMV2 (2M CBR) / WMA2 128k | 1280x720 | 10.09s | 2.68 MB |
| test_1080p_asf_wmv2.wmv | ASF / WMV2 (4M CBR) / WMA2 128k | 1920x1080 | 10.09s | 5.06 MB |
| test_720p_mpg_mpeg2.mpg | MPEG-PS / MPEG-2 (4M) / MP2 192k | 1280x720 | 10.01s | 5.08 MB |
| test_1080p_mpg_mpeg2.mpg | MPEG-PS / MPEG-2 (8M) / MP2 192k | 1920x1080 | 10.01s | 9.83 MB |
| test_720p_3gp_mpeg4.3gp | 3GP / MPEG-4 (1.5M) / AAC 64k | 1280x720 | 10.00s | 2.03 MB |
| test_1080p_3gp_mpeg4.3gp | 3GP / MPEG-4 (2.5M) / AAC 64k | 1920x1080 | 10.00s | 3.90 MB |
| test_720p_ogv_theora.ogv | Ogg / Theora (q6) / Vorbis q4 | 1280x720 | 10.00s | 5.52 MB |
| test_1080p_ogv_theora.ogv | Ogg / Theora (q6) / Vorbis q4 | **1920x1072** | 10.00s | 10.20 MB |

全部 30 个文件已通过解码自检（`ffmpeg -v error -i <file> -f null -`，无报错）且时长解析为 10 秒。

## 注意事项与实现取舍

1. **3GP 用 MPEG-4 而非 H.263**：H.263 编码器最高只支持 CIF(352x288)，无法编码 720p/1080p，故改用 3GP 标准同样支持的 MPEG-4 + AAC。
2. **Theora 的 1080p 实际是 1920x1072**：libtheora 要求宽高均为 16 的倍数，脚本用 `scale=trunc(iw/16)*16:trunc(ih/16)*16` 处理，并通过 SAR 134:135 保持 16:9 显示比例；720p（1280x720）本身已满足，不受影响。
3. **ProRes 体积**：ProRes 为帧内编码，默认使用 `profile 0 (Proxy)` 控制体积（1080p 10 秒约 28 MB / 22 Mbps）。如需更高质量，把脚本中 `mov_prores` 的 `-profile:v 0` 改成 `1`(LT)/`2`(Standard)/`3`(HQ)，体积会成倍增长。
4. **Opus 必须 48kHz**：libopus 编码器只接受 48000Hz，脚本对 `mkv_vp9` / `webm_vp9` 单独指定 `-ar 48000`，其余格式统一 44100Hz。
5. **AVI / FLV 用 MP3**：AVI 内封装 AAC 兼容性差，FLV 场景传统上使用 MP3，故这两个组合使用 libmp3lame。
6. **WMV2 不支持 CRF**：使用显式码率（720p 2M / 1080p 4M）。
7. **MPEG-PS 码率敏感**：`mpg_mpeg2` 若报 buffer underflow，脚本会自动把码率减半重试一次。
8. **VP9/VP8 编码速度**：统一加 `-deadline good -cpu-used 4 -row-mt 1`，1080p 10 秒约 5 秒完成。

## 目录内其他文件

- `generate_test_videos.ps1`：生成/校验脚本（格式清单位于脚本中部 `$FormatList`，可直接增删条目扩展更多格式）。
