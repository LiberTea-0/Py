import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import make_interp_spline

# 定义关键点
x_points = [0, 2, 4, 6, 8, 10, 12]  # x 轴
y_points = [2.5, 2.25, 2.25, 2.25, 2.2, 2.2, 2.2]  # y 轴

# 将关键点分为两组
x_first_press = x_points[:4]
y_first_press = y_points[:4]

x_second_press = x_points[3:]
y_second_press = y_points[3:]

# 分别生成插值平滑曲线
x_dense_first = np.linspace(x_first_press[0], x_first_press[-1], 150)
y_smooth_first = make_interp_spline(x_first_press, y_first_press, k=2)(x_dense_first)

x_dense_second = np.linspace(x_second_press[0], x_second_press[-1], 300)
y_smooth_second = make_interp_spline(x_second_press, y_second_press, k=2)(x_dense_second)

# 添加噪声
noise_amplitude = 0.05
noise_first = np.random.normal(0, noise_amplitude, size=y_smooth_first.shape)
y_noisy_first = y_smooth_first + noise_first

noise_second = np.random.normal(0, noise_amplitude, size=y_smooth_second.shape)
y_noisy_second = y_smooth_second + noise_second

# 绘图
fig, ax = plt.subplots(figsize=(10, 5))

# 绘制第一次按压曲线
ax.plot(
    x_dense_first,
    y_noisy_first,
    color="blue",
    linewidth=1,
    label="First Press",
)

# 绘制第二次按压曲线
ax.plot(
    x_dense_second,
    y_noisy_second,
    color="green",
    linewidth=1,
    label="Second Press",
)

# 设置图形样式
ax.set_title("Press Events", fontsize=16)
ax.set_xlabel("Time / s", fontsize=12)
ax.set_ylabel("Voltage / V", fontsize=12)
ax.grid(True, linestyle="--", alpha=0.6)
ax.legend()

# 显示图形
plt.show()
