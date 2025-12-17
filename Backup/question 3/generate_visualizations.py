"""
Script to generate all visualizations from the News Classification project
and save them as image files for insertion into the Word document
"""

import os
import re
import string
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.model_selection import train_test_split
from sklearn.metrics import classification_report, confusion_matrix, f1_score, accuracy_score
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.decomposition import PCA, TruncatedSVD
from sklearn.cluster import KMeans
from sklearn.linear_model import LogisticRegression
import tensorflow as tf
from tensorflow.keras.preprocessing.text import Tokenizer
from tensorflow.keras.preprocessing.sequence import pad_sequences
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import Embedding, Conv1D, GlobalMaxPooling1D, Dense, Dropout
from sklearn.preprocessing import LabelEncoder
from tensorflow.keras.utils import to_categorical

# Set style
sns.set(style="whitegrid")
plt.rcParams["figure.figsize"] = (10, 6)
plt.rcParams['figure.dpi'] = 300  # High resolution for report

# Set random seeds
RANDOM_STATE = 42
np.random.seed(RANDOM_STATE)
tf.random.set_seed(RANDOM_STATE)

# Create output directory for images
output_dir = os.path.join(os.path.dirname(__file__), 'report_images')
os.makedirs(output_dir, exist_ok=True)

print("=" * 60)
print("Generating Visualizations for Report")
print("=" * 60)

# ==================== Load Data ====================
print("\n1. Loading dataset...")
DATA_DIR = r"C:\INT_SYSTEMS\Final exam\FINAL_EXAM_INTELLIGENT_SYSTEM\qtn_3\review\News Articles"

categories = []
texts = []
filenames = []

for category in os.listdir(DATA_DIR):
    category_path = os.path.join(DATA_DIR, category)
    if not os.path.isdir(category_path):
        continue
    for fname in os.listdir(category_path):
        fpath = os.path.join(category_path, fname)
        if not os.path.isfile(fpath):
            continue
        with open(fpath, "r", encoding="latin-1") as f:
            text = f.read().strip()
        categories.append(category)
        texts.append(text)
        filenames.append(fname)

df = pd.DataFrame({
    "category": categories,
    "text": texts,
    "filename": filenames
})

print(f"   Loaded {len(df)} articles")

# ==================== Clean Text ====================
print("\n2. Cleaning text data...")
def clean_text(text):
    text = text.lower()
    text = re.sub(r"<.*?>", " ", text)
    text = re.sub(r"\d+", " ", text)
    text = text.translate(str.maketrans("", "", string.punctuation))
    text = re.sub(r"\s+", " ", text).strip()
    return text

df["clean_text"] = df["text"].apply(clean_text)

# ==================== Create Features ====================
print("\n3. Creating features...")
tfidf_vectorizer = TfidfVectorizer(
    max_features=5000,
    ngram_range=(1, 2),
    stop_words="english"
)
X_tfidf = tfidf_vectorizer.fit_transform(df["clean_text"])
y = df["category"]

MAX_NUM_WORDS = 10000
MAX_SEQ_LEN = 300
tokenizer = Tokenizer(num_words=MAX_NUM_WORDS, oov_token="<OOV>")
tokenizer.fit_on_texts(df["clean_text"])
sequences = tokenizer.texts_to_sequences(df["clean_text"])
X_seq = pad_sequences(sequences, maxlen=MAX_SEQ_LEN, padding="post", truncating="post")

# ==================== Train/Test Split ====================
print("\n4. Splitting data...")
label_encoder = LabelEncoder()
y_int = label_encoder.fit_transform(y)
num_classes = len(label_encoder.classes_)

indices = np.arange(len(y_int))
X_tfidf_train, X_tfidf_test, X_seq_train, X_seq_test, y_train_int, y_test_int, train_idx, test_idx = train_test_split(
    X_tfidf, X_seq, y_int, indices,
    test_size=0.2,
    random_state=RANDOM_STATE,
    stratify=y_int
)

y_train_cat = to_categorical(y_train_int, num_classes=num_classes)
y_test_cat = to_categorical(y_test_int, num_classes=num_classes)

# ==================== Visualization 1: Class Distribution ====================
print("\n5. Generating Class Distribution chart...")
class_counts = df["category"].value_counts().sort_index()

plt.figure(figsize=(10, 6))
sns.barplot(x=class_counts.index, y=class_counts.values, palette="viridis")
plt.title("Class Distribution of News Articles", fontsize=14, fontweight='bold')
plt.xlabel("Category", fontsize=12)
plt.ylabel("Count", fontsize=12)
plt.xticks(rotation=45)
plt.grid(axis='y', alpha=0.3)
for i, v in enumerate(class_counts.values):
    plt.text(i, v + 5, str(v), ha='center', va='bottom', fontweight='bold')
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '1_class_distribution.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 1_class_distribution.png")

# ==================== Train Models ====================
print("\n6. Training ML model...")
ml_model = LogisticRegression(max_iter=2000, random_state=RANDOM_STATE)
ml_model.fit(X_tfidf_train, y_train_int)
y_ml_pred = ml_model.predict(X_tfidf_test)
y_ml_prob = ml_model.predict_proba(X_tfidf_test)

print("\n7. Training DL model...")
EMBEDDING_DIM = 100
dl_model = Sequential([
    Embedding(input_dim=MAX_NUM_WORDS, output_dim=EMBEDDING_DIM),
    Conv1D(filters=128, kernel_size=5, activation="relu"),
    GlobalMaxPooling1D(),
    Dropout(0.5),
    Dense(64, activation="relu"),
    Dropout(0.5),
    Dense(num_classes, activation="softmax")
])

dl_model.compile(
    loss="categorical_crossentropy",
    optimizer="adam",
    metrics=["accuracy"]
)

X_seq_train_final, X_seq_val, y_train_cat_final, y_val_cat = train_test_split(
    X_seq_train, y_train_cat,
    test_size=0.2,
    random_state=RANDOM_STATE,
    stratify=np.argmax(y_train_cat, axis=1)
)

EPOCHS = 20
BATCH_SIZE = 32

print("   Training CNN (this may take a few minutes)...")
history = dl_model.fit(
    X_seq_train_final, y_train_cat_final,
    validation_data=(X_seq_val, y_val_cat),
    epochs=EPOCHS,
    batch_size=BATCH_SIZE,
    verbose=0
)

y_dl_prob = dl_model.predict(X_seq_test, verbose=0)
y_dl_pred = np.argmax(y_dl_prob, axis=1)

# ==================== Visualization 2: ML Confusion Matrix ====================
print("\n8. Generating ML Confusion Matrix...")
cm_ml = confusion_matrix(y_test_int, y_ml_pred)
plt.figure(figsize=(10, 8))
sns.heatmap(cm_ml, annot=True, fmt="d", cmap="Blues",
            xticklabels=label_encoder.classes_,
            yticklabels=label_encoder.classes_)
plt.title("Confusion Matrix – Classical ML Model (Logistic Regression)", fontsize=14, fontweight='bold')
plt.xlabel("Predicted Label", fontsize=12)
plt.ylabel("True Label", fontsize=12)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '2_ml_confusion_matrix.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 2_ml_confusion_matrix.png")

# ==================== Visualization 3: DL Training Curves ====================
print("\n9. Generating DL Training Curves...")
plt.figure(figsize=(12, 5))

plt.subplot(1, 2, 1)
plt.plot(history.history['accuracy'], label='Training Accuracy', marker='o')
plt.plot(history.history['val_accuracy'], label='Validation Accuracy', marker='s')
plt.title('Model Accuracy', fontsize=14, fontweight='bold')
plt.xlabel('Epoch', fontsize=12)
plt.ylabel('Accuracy', fontsize=12)
plt.legend()
plt.grid(True, alpha=0.3)
plt.ylim([0, 1])

plt.subplot(1, 2, 2)
plt.plot(history.history['loss'], label='Training Loss', marker='o')
plt.plot(history.history['val_loss'], label='Validation Loss', marker='s')
plt.title('Model Loss', fontsize=14, fontweight='bold')
plt.xlabel('Epoch', fontsize=12)
plt.ylabel('Loss', fontsize=12)
plt.legend()
plt.grid(True, alpha=0.3)

plt.tight_layout()
plt.savefig(os.path.join(output_dir, '3_dl_training_curves.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 3_dl_training_curves.png")

# ==================== Visualization 4: DL Confusion Matrix ====================
print("\n10. Generating DL Confusion Matrix...")
cm_dl = confusion_matrix(y_test_int, y_dl_pred)
plt.figure(figsize=(10, 8))
sns.heatmap(cm_dl, annot=True, fmt="d", cmap="Greens",
            xticklabels=label_encoder.classes_,
            yticklabels=label_encoder.classes_)
plt.title("Confusion Matrix – Deep Learning Model (CNN)", fontsize=14, fontweight='bold')
plt.xlabel("Predicted Label", fontsize=12)
plt.ylabel("True Label", fontsize=12)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '4_dl_confusion_matrix.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 4_dl_confusion_matrix.png")

# ==================== Clustering ====================
print("\n11. Performing clustering...")
n_components = 100
svd = TruncatedSVD(n_components=n_components, random_state=RANDOM_STATE)
X_tfidf_reduced = svd.fit_transform(X_tfidf)
X_clustering = X_tfidf_reduced

NUM_CLUSTERS = 5
kmeans = KMeans(n_clusters=NUM_CLUSTERS, random_state=RANDOM_STATE, n_init=10)
clusters = kmeans.fit_predict(X_clustering)
df["cluster"] = clusters

# ==================== Visualization 5: Clusters Visualization ====================
print("\n12. Generating Clusters Visualization...")
pca = PCA(n_components=2, random_state=RANDOM_STATE)
X_pca = pca.fit_transform(X_clustering)

plt.figure(figsize=(12, 8))
scatter = plt.scatter(X_pca[:, 0], X_pca[:, 1], c=clusters, cmap='viridis', alpha=0.6, s=50)
plt.colorbar(scatter, label='Cluster ID')
plt.title("K-Means Clusters Visualization (PCA Projection)", fontsize=14, fontweight='bold')
plt.xlabel(f"PC1 (Explained Variance: {pca.explained_variance_ratio_[0]:.2%})", fontsize=12)
plt.ylabel(f"PC2 (Explained Variance: {pca.explained_variance_ratio_[1]:.2%})", fontsize=12)
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '5_clusters_visualization.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 5_clusters_visualization.png")

# ==================== RL Agent ====================
print("\n13. Setting up RL environment...")
ml_max_conf = y_ml_prob.max(axis=1)
dl_max_conf = y_dl_prob.max(axis=1)

X_tfidf_test_reduced = svd.transform(X_tfidf_test)
clusters_test = kmeans.predict(X_tfidf_test_reduced)

lengths = np.sum(X_seq_test != 0, axis=1)
def bin_length(l):
    if l < 100:
        return 0
    elif l < 200:
        return 1
    else:
        return 2
length_bins = np.array([bin_length(l) for l in lengths])
disagreement = (y_ml_pred != y_dl_pred).astype(int)
true_labels = y_test_int

def bin_conf(p):
    if p < 0.2:
        return 0
    elif p < 0.4:
        return 1
    elif p < 0.6:
        return 2
    elif p < 0.8:
        return 3
    else:
        return 4

ml_conf_bins = np.array([bin_conf(p) for p in ml_max_conf])
dl_conf_bins = np.array([bin_conf(p) for p in dl_max_conf])

num_conf_bins = 5
num_length_bins = 3
num_clusters = NUM_CLUSTERS
num_disagree = 2
num_actions = 3

def encode_state(i):
    return (
        ml_conf_bins[i] * (num_conf_bins * num_length_bins * num_clusters * num_disagree) +
        dl_conf_bins[i] * (num_length_bins * num_clusters * num_disagree) +
        length_bins[i] * (num_clusters * num_disagree) +
        clusters_test[i] * num_disagree +
        disagreement[i]
    )

def get_reward(i, action):
    true = true_labels[i]
    ml_correct = (y_ml_pred[i] == true)
    dl_correct = (y_dl_pred[i] == true)
    ml_conf = ml_max_conf[i]
    dl_conf = dl_max_conf[i]
    
    if action == 0:
        if ml_correct:
            reward = 5 + 2 * ml_conf
        else:
            reward = -5
    elif action == 1:
        if dl_correct:
            reward = 6 + 2 * dl_conf
        else:
            reward = -6
    else:
        if (not ml_correct) and (not dl_correct):
            reward = 2
        else:
            reward = -1
    return reward

# Train Q-Learning
print("\n14. Training Q-Learning agent...")
num_states = num_conf_bins**2 * num_length_bins * num_clusters * num_disagree
Q = np.zeros((num_states, num_actions))
num_episodes = 1200
alpha = 0.1
gamma = 0.9
epsilon = 0.2

rewards_per_episode = []
n_samples = len(true_labels)

for ep in range(num_episodes):
    total_reward = 0.0
    indices = np.random.permutation(n_samples)
    
    for idx in indices:
        s = encode_state(idx)
        if np.random.rand() < epsilon:
            a = np.random.randint(num_actions)
        else:
            a = np.argmax(Q[s])
        
        r = get_reward(idx, a)
        total_reward += r
        
        s_next = s
        Q[s, a] = Q[s, a] + alpha * (r + gamma * np.max(Q[s_next]) - Q[s, a])
    
    rewards_per_episode.append(total_reward)
    
    if (ep + 1) % 200 == 0:
        print(f"   Episode {ep + 1}/{num_episodes} - Average reward: {np.mean(rewards_per_episode[-200:]):.2f}")

# ==================== Visualization 6: Q-Learning Training Progress ====================
print("\n15. Generating Q-Learning Training Progress...")
plt.figure(figsize=(12, 5))
plt.plot(rewards_per_episode, alpha=0.6, linewidth=0.5, label='Episode Reward')
plt.plot(pd.Series(rewards_per_episode).rolling(50).mean(), linewidth=2, label='Moving Average (50 episodes)')
plt.title("Q-Learning Training Progress", fontsize=14, fontweight='bold')
plt.xlabel("Episode", fontsize=12)
plt.ylabel("Total Reward per Episode", fontsize=12)
plt.legend()
plt.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '6_qlearning_progress.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 6_qlearning_progress.png")

# ==================== Visualization 7: RL Action Distribution ====================
print("\n16. Generating RL Action Distribution...")
rl_actions = []
for i in range(n_samples):
    s = encode_state(i)
    a = np.argmax(Q[s])
    rl_actions.append(a)

rl_actions = np.array(rl_actions)
action_counts = [np.sum(rl_actions==0), np.sum(rl_actions==1), np.sum(rl_actions==2)]

plt.figure(figsize=(10, 6))
plt.bar(["ML", "DL", "Escalate"], action_counts, color=['skyblue', 'lightgreen', 'salmon'])
plt.title("RL Action Distribution on Test Set", fontsize=14, fontweight='bold')
plt.ylabel("Count", fontsize=12)
plt.xlabel("Action", fontsize=12)
for i, v in enumerate(action_counts):
    plt.text(i, v + 5, str(v), ha='center', va='bottom', fontweight='bold')
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(output_dir, '7_rl_action_distribution.png'), dpi=300, bbox_inches='tight')
plt.close()
print("   Saved: 7_rl_action_distribution.png")

print("\n" + "=" * 60)
print("All visualizations generated successfully!")
print(f"Images saved in: {output_dir}")
print("=" * 60)

