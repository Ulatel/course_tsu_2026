import os
import sys
import time
import tempfile
import traceback

sys.stdout.reconfigure(encoding="utf-8")
import matplotlib
matplotlib.use("Agg")

PLOT_DIR = os.path.dirname(os.path.abspath(__file__))

results = {}

def run(name, fn):
    t0 = time.time()
    try:
        fn()
        results[name] = f"PASS ({time.time()-t0:.1f}s)"
    except Exception as e:
        results[name] = f"FAIL: {type(e).__name__}: {e}"
        traceback.print_exc()


def test_m1():
    import numpy as np
    import pandas as pd
    import matplotlib.pyplot as plt
    import seaborn as sns
    from sklearn.preprocessing import LabelEncoder, StandardScaler, MinMaxScaler
    from sklearn.model_selection import train_test_split, cross_val_score
    from sklearn.neighbors import KNeighborsClassifier
    from sklearn.metrics import accuracy_score
    from sklearn.impute import SimpleImputer

    df = pd.DataFrame({
        "age": [22, 38, 26, 35, 45, np.nan, 30, 40],
        "sex": ["male", "female", "female", "male", "male", "male", "female", "male"],
        "pclass": [3, 1, 3, 1, 2, 3, 2, 1],
        "survived": [0, 1, 1, 1, 0, 0, 1, 0],
    })
    df["sex_enc"] = LabelEncoder().fit_transform(df["sex"])
    df["age"] = SimpleImputer(strategy="median").fit_transform(df[["age"]])
    X = StandardScaler().fit_transform(df[["age", "pclass", "sex_enc"]])
    X_tr, X_te, y_tr, y_te = train_test_split(X, df["survived"], test_size=0.25, random_state=42)
    knn = KNeighborsClassifier(n_neighbors=3).fit(X_tr, y_tr)
    acc = accuracy_score(y_te, knn.predict(X_te))
    assert 0 <= acc <= 1
    cross_val_score(knn, X, df["survived"], cv=3)
    sns.histplot(df["age"])
    plt.savefig(os.path.join(PLOT_DIR, "m1_plot.png"))
    plt.close()


def test_m2():
    import numpy as np
    from sklearn.linear_model import LinearRegression, Ridge, Lasso, SGDRegressor
    from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
    from sklearn.model_selection import GridSearchCV, cross_val_score
    from sklearn.preprocessing import StandardScaler

    rng = np.random.RandomState(0)
    X = rng.rand(200, 3)
    y = X @ np.array([3.0, -2.0, 1.0]) + 5.0 + rng.randn(200) * 0.1

    # градиентный спуск "с нуля"
    Xb = np.hstack([np.ones((200, 1)), X])
    w = np.zeros(4)
    for _ in range(500):
        grad = 2 / 200 * Xb.T @ (Xb @ w - y)
        w -= 0.1 * grad
    assert np.allclose(w, [5.0, 3.0, -2.0, 1.0], atol=0.05)

    lr = LinearRegression().fit(X, y)
    assert r2_score(y, lr.predict(X)) > 0.99
    Ridge(alpha=1.0).fit(X, y)
    Lasso(alpha=0.1).fit(X, y)
    SGDRegressor(max_iter=1000, tol=1e-3).fit(StandardScaler().fit_transform(X), y)
    gs = GridSearchCV(Ridge(), {"alpha": [0.1, 1.0, 10.0]}, cv=3).fit(X, y)
    assert gs.best_params_
    assert mean_absolute_error(y, lr.predict(X)) < 0.5
    assert mean_squared_error(y, lr.predict(X)) > 0
    cross_val_score(lr, X, y, cv=3, scoring="r2")


def test_m3():
    import numpy as np
    import xgboost as xgb
    import lightgbm as lgb
    from sklearn.ensemble import RandomForestClassifier
    from sklearn.metrics import roc_auc_score
    from imblearn.over_sampling import SMOTE

    rng = np.random.RandomState(0)
    n = 400
    X = rng.rand(n, 10)
    y = (X[:, 0] + X[:, 1] > 1.2).astype(int)
    y[rng.choice(n, 30, replace=False)] = 1  # дисбаланс

    X_res, y_res = SMOTE(random_state=42).fit_resample(X, y)
    assert len(y_res) > n

    for model in (xgb.XGBClassifier(n_estimators=20, eval_metric="logloss"),
                  lgb.LGBMClassifier(n_estimators=20, verbose=-1),
                  RandomForestClassifier(n_estimators=20)):
        model.fit(X_res, y_res)
        auc = roc_auc_score(y, model.predict_proba(X)[:, 1])
        assert auc > 0.8


def test_m4():
    import torch
    import torch.nn as nn
    import torch.nn.functional as F
    from torch.utils.data import DataLoader, TensorDataset
    from torchvision import datasets, transforms

    torch.manual_seed(0)
    # MNIST-подобный перцептрон на синтетике (датасеты torchvision доступны)
    assert hasattr(datasets, "MNIST")
    tfm = transforms.Compose([transforms.ToTensor(), transforms.Normalize((0.5,), (0.5,))])
    assert tfm is not None

    X = torch.rand(128, 1, 28, 28)
    y = torch.randint(0, 10, (128,))

    class MLP(nn.Module):
        def __init__(self):
            super().__init__()
            self.fc1 = nn.Linear(28 * 28, 64)
            self.fc2 = nn.Linear(64, 10)

        def forward(self, x):
            return self.fc2(F.relu(self.fc1(x.flatten(1))))

    model = MLP()
    opt = torch.optim.Adam(model.parameters(), lr=1e-3)
    loader = DataLoader(TensorDataset(X, y), batch_size=32)
    for xb, yb in loader:
        loss = F.cross_entropy(model(xb), yb)
        opt.zero_grad(); loss.backward(); opt.step()
    assert loss.item() > 0


def test_m5():
    import torch
    import torch.nn as nn

    # PINN: du/dx = cos(x), u(0)=0  =>  u(x) = sin(x), физический residual в loss
    torch.manual_seed(0)
    x = torch.linspace(0, 2, 50, requires_grad=True).reshape(-1, 1)
    model = nn.Sequential(nn.Linear(1, 32), nn.Tanh(), nn.Linear(32, 1))
    opt = torch.optim.Adam(model.parameters(), lr=1e-2)
    for _ in range(300):
        u = model(x)
        du_dx = torch.autograd.grad(u.sum(), x, create_graph=True)[0]
        phys = ((du_dx - torch.cos(x)) ** 2).mean()
        bc = (model(torch.zeros(1, 1)) - 0.0) ** 2
        loss = phys + bc
        opt.zero_grad(); loss.backward(); opt.step()
    assert loss.item() < 1e-2, f"PINN loss too high: {loss.item()}"
    with torch.no_grad():
        err = (model(x) - torch.sin(x)).abs().max().item()
    assert err < 0.1, f"max err {err}"


def test_m6():
    import numpy as np
    import albumentations as A
    import ultralytics
    from ultralytics import YOLO
    import label_studio
    from torchvision import transforms as tvt

    assert ultralytics.__version__ and label_studio.__version__

    # аугментация CV
    img = (np.random.rand(64, 64, 3) * 255).astype("uint8")
    aug = A.Compose([A.HorizontalFlip(p=1.0), A.RandomBrightnessContrast(p=1.0)])
    out = aug(image=img)["image"]
    assert out.shape == (64, 64, 3)

    # UNet-подобная сегментация (torchvision) — Down/Up через transforms не нужен,
    # проверяем building blocks
    from torchvision.models.segmentation import fcn_resnet50
    import torch
    m = fcn_resnet50(weights=None, num_classes=2)
    out = m(torch.rand(1, 3, 64, 64))["out"]
    assert out.shape == (1, 2, 64, 64)

    # YOLO: крошечная модель (загрузка весов yolo11n.pt ~5МБ)
    model = YOLO("yolo11n.pt")
    res = model.predict(source=np.zeros((64, 64, 3), dtype="uint8"), verbose=False)
    assert res is not None


def test_m7():
    from sklearn.pipeline import Pipeline
    from sklearn.preprocessing import StandardScaler
    from sklearn.impute import SimpleImputer
    from sklearn.linear_model import LogisticRegression
    from sklearn.model_selection import cross_val_score
    import numpy as np

    rng = np.random.RandomState(0)
    X = rng.rand(100, 5)
    y = (X.sum(axis=1) > 2.5).astype(int)
    pipe = Pipeline([
        ("impute", SimpleImputer()),
        ("scale", StandardScaler()),
        ("clf", LogisticRegression()),
    ])
    assert cross_val_score(pipe, X, y, cv=3).mean() > 0.8


def test_m8():
    import numpy as np
    from transformers import AutoTokenizer, AutoModel
    import sentence_transformers
    from sklearn.cluster import KMeans

    # кластеризация текстов энкодером (маленькая модель ~30МБ)
    from sentence_transformers import SentenceTransformer
    model = SentenceTransformer("paraphrase-MiniLM-L3-v2")
    texts = ["физика частиц", "нейронные сети", "бозон хиггса", "обучение модели"]
    emb = model.encode(texts, normalize_embeddings=True)
    assert emb.shape == (4, 384)
    labels = KMeans(n_clusters=2, n_init=10, random_state=0).fit_predict(emb)
    assert len(set(labels)) == 2


def test_extra():
    # доп. стек: scipy, statsmodels, catboost, opencv, pycocotools, HF fine-tuning
    import numpy as np
    from scipy import stats as sps
    import statsmodels.api as sm
    from catboost import CatBoostClassifier
    import cv2
    import PIL
    import pycocotools
    from pycocotools import mask as cocomask
    import datasets, accelerate, peft, tqdm

    # scipy: t-тест
    a, b = np.random.rand(100), np.random.rand(100) + 0.5
    assert sps.ttest_ind(a, b).pvalue < 1e-6

    # statsmodels: OLS с доверительными интервалами (М2)
    X = np.random.rand(100, 2)
    y = X @ [2.0, -1.0] + 1.0 + np.random.rand(100) * 0.01
    ols = sm.OLS(y, sm.add_constant(X)).fit()
    assert ols.rsquared > 0.99

    # catboost (М3)
    rng = np.random.RandomState(0)
    Xc, yc = rng.rand(200, 5), (rng.rand(200) > 0.5).astype(int)
    acc = CatBoostClassifier(iterations=10, verbose=0).fit(Xc, yc).score(Xc, yc)
    assert acc > 0.8

    # opencv + Pillow (М6)
    img = (np.random.rand(64, 64, 3) * 255).astype("uint8")
    assert cv2.cvtColor(img, cv2.COLOR_RGB2GRAY).shape == (64, 64)
    assert PIL.Image.fromarray(img).size == (64, 64)

    # pycocotools: COCO-маски (М6, конвертация разметки)
    m = cocomask.encode(np.asfortranarray((np.random.rand(32, 32) > 0.5).astype("uint8")))
    assert cocomask.area(m) >= 0


run("М1  EDA+KNN (Titanic)", test_m1)
run("М2  Регрессия+GD+Ridge/Lasso+GridSearch", test_m2)
run("М3  XGBoost+LightGBM+RF+SMOTE", test_m3)
run("М4  PyTorch MLP (MNIST-style)", test_m4)
run("М5  PINN (физический loss)", test_m5)
run("М6  YOLO+albumentations+UNet+label-studio", test_m6)
run("М7  Pipeline", test_m7)
run("М8  sentence-transformers+KMeans", test_m8)
run("Доп  scipy+statsmodels+catboost+opencv+coco+HF", test_extra)

print("\n========== ИТОГ ==========")
for k, v in results.items():
    print(f"{k}: {v}")
fails = [k for k, v in results.items() if v.startswith("FAIL")]
print(f"\nПройдено: {len(results)-len(fails)}/{len(results)}")
sys.exit(1 if fails else 0)
