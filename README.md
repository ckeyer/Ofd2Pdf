# Ofd2Pdf

Convert OFD files to PDF files.

## GUI（Windows）

![Ofd2Pdf GUI](screenshot.png)

你可以单击打开 GUI，选择文件或者拖拽文件到列表，亦或者简单将文件拖拽到可执行文件上，它将自动完成转换过程。

## HTTP API 服务（Linux / .NET 8）

`Ofd2PdfService` 是一个基于 ASP.NET Core 8 的 HTTP 接口服务，可运行在 Linux 服务器上，通过接口将 OFD 文件转换为 PDF。

### 快速启动

#### 方式一：直接运行

```bash
cd Ofd2PdfService
dotnet run
```

服务默认监听 `http://localhost:5000`。

#### 方式二：Docker 容器

```bash
cd Ofd2PdfService

# 构建镜像
docker build -t ofd2pdf-service .

# 运行容器（监听宿主机 8080 端口）
docker run -p 8080:8080 ofd2pdf-service
```

### 接口说明

#### `POST /api/convert`

上传 OFD 文件，返回转换后的 PDF 文件。

| 参数 | 类型 | 说明 |
|------|------|------|
| `file` | `multipart/form-data` | 待转换的 `.ofd` 文件 |

**示例（curl）：**

```bash
curl -X POST http://localhost:8080/api/convert \
     -F "file=@sample.ofd" \
     --output result.pdf
```

**成功响应：** `200 OK`，`Content-Type: application/pdf`，响应体为 PDF 文件内容。

**错误响应：**

| 状态码 | 原因 |
|--------|------|
| `400 Bad Request` | 未上传文件或文件不是 `.ofd` 格式 |
| `500 Internal Server Error` | 转换过程中发生错误 |
