# 使用官方的基础镜像
FROM ghcr.io/astral-sh/uv:python3.12-alpine

# 添加项目代码
ADD . /CalibrationTableGeneration
WORKDIR /CalibrationTableGeneration

# 安装依赖
RUN uv sync --locked

EXPOSE 5000

# 运行应用
CMD ["uv", "run", "app.py"]