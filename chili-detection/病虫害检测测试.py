from ultralytics import YOLO
import os


save_project = r"D:\指天椒检测\指天椒检测\病虫害图片\bch检测数据"

os.makedirs(save_project, exist_ok=True)

model = YOLO(r"D:\指天椒检测\指天椒检测\训练好的模型\detect\train11\weights\best.pt")

# 预测并指定保存路径
results = model.predict(
    source=r"D:\指天椒检测\指天椒检测\病虫害图片\bch图片\白粉虱.jpg",
    save=True,
    show=False,
    save_txt=True,
    project=save_project,
    name="predict",
    exist_ok=True
)

final_save_path = os.path.join(save_project, "predict")
print("预测完成！")
print(f"预测结果保存到: {final_save_path}")