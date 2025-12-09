# MeowMail 部署指南

本文档将指导你如何从源码部署 **MeowMail**，包括克隆项目、修改配置以及通过 Docker Compose 启动服务。

---

## 📦 克隆项目

```bash
git clone https://github.com/defeatd/meowmail.git
cd meowmail
```

---

## ⚙️ 修改配置文件

使用编辑器（例如 nano）打开 `docker-compose.yml`：

```bash
nano docker-compose.yml
```

在文件中你可以根据需要修改：

* **端口映射**（如 `80:80` 改成你想使用的端口）
* **环境变量秘钥**（JWT_SECRET）

保存并退出：

* 按下 `Ctrl + O` 保存
* 按下 `Ctrl + X` 退出

---

## ▶️ 构建并启动服务

运行以下命令：

```bash
docker compose up --build -d
```

这会在后台启动服务容器并完成构建。

---

## ✅ 启动完成

部署成功后，你可以通过浏览器访问：

```
http://服务器IP:你设置的端口
```

---

## 🔧 改进说明

相较原版项目，本版本进行了以下两点优化：

1. **优化收件逻辑**：在处理邮件时，会先将垃圾邮件文件夹中的所有邮件移动到收件箱，再统一从收件箱获取邮件，确保不会遗漏关键信息。
2. **本地构建镜像**：部署时改为基于本地源码构建 Docker 镜像，而不再从远端仓库拉取，提高了可控性与构建速度。

## 📄 其他说明

如需停止服务：

```bash
docker compose down
```

如需查看日志：

```bash
docker compose logs -f
```

欢迎为项目贡献代码或提交 issue！
