from docx import Document
import string
import re
import nltk
from nltk.tokenize import word_tokenize
from nltk.probability import FreqDist
from nltk.corpus import stopwords
from natasha import MorphVocab, NewsEmbedding, NewsMorphTagger, Doc as NatashaDoc, Segmenter
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# === 1. Загрузка Word-документа ===
doc = Document(r"c:\Users\HYPERPC\Documents\Научная статья\Гуценко Павел. реферат.docx")
text = "\n".join([para.text for para in doc.paragraphs])

# === 2. Очистка текста ===
text = text.lower()
spec_chars = string.punctuation + '«»\t—…’'
text = "".join([ch for ch in text if ch not in spec_chars])
text = re.sub('\n', ' ', text)
text = "".join([ch for ch in text if ch not in string.digits])

# === 3. Токенизация NLTK ===
nltk.download('punkt')
nltk.download('stopwords')
tokens = word_tokenize(text)

# === 4. Удаление стоп-слов ===
stop_words = set(stopwords.words('russian'))

custom_exclude = {
    'вопрос', 'это', 'эти', 'также', 'однако', 'поэтому', 'вообще', 'например',
    'конечно', 'безусловно', 'вроде', 'впрочем', 'действительно', 'вероятно',
    'итак', 'причем', 'собственно', 'между', 'просто', 'да', 'нет', 'ли', 'же',
    'бы', 'б', 'пусть', 'ведь', 'хотя', 'если', 'чтобы', 'и', 'а', 'но', 'или',
    'что', 'как', 'когда', 'где', 'куда', 'откуда', 'с', 'со', 'в', 'во', 'на', 'над', 'под',
    'к', 'ко', 'по', 'у', 'о', 'об', 'про', 'для', 'из', 'за', 'при',
    'я', 'ты', 'он', 'она', 'оно', 'мы', 'вы', 'они', 'мой', 'твой', 'его', 'ее', 'наш', 'ваш', 'их',
    'меня', 'тебя', 'него', 'нее', 'нас', 'вас', 'них', 'мне', 'тебе', 'ему', 'ей', 'нам', 'вам', 'им',
    'мной', 'тобой', 'ним', 'ней', 'нами', 'вами', 'ими', 'обоих', 'набиуллин', 'касаться', 'видеть', 'ключевой', 'денежнокредитный',
    'говорить', 'считать', 'сказать', 'который',
}

tokens_no_stopwords = [
    word for word in tokens
    if word not in stop_words
    and word not in custom_exclude
    and word.isalpha()
]

# === 5. Лемматизация и POS Natasha ===
morph_vocab = MorphVocab()
emb = NewsEmbedding()
morph_tagger = NewsMorphTagger(emb)
segmenter = Segmenter()

natasha_doc = NatashaDoc(" ".join(tokens_no_stopwords))
natasha_doc.segment(segmenter)
natasha_doc.tag_morph(morph_tagger)

for token in natasha_doc.tokens:
    token.lemmatize(morph_vocab)

# === 6. Фильтрация и частотный анализ ===
lemmas = [
    token.lemma for token in natasha_doc.tokens
    if token.text.isalpha()
    and token.lemma not in custom_exclude
]

fdist = FreqDist(lemmas)
print("\n🔝 Топ-10 частотных слов:")
for word, freq in fdist.most_common(10):
    print((word, freq))

count_zestkiy = sum(1 for token in natasha_doc.tokens if token.lemma == "жесткий")
print("\nКоличество появлений слова 'жесткий' (лемма):", count_zestkiy)

# === 7. Выделение глаголов и прилагательных ===
verbs = [
    token.lemma for token in natasha_doc.tokens
    if token.pos in {'VERB', 'INFN'}
    and token.lemma not in custom_exclude
]

adjectives = [
    token.lemma for token in natasha_doc.tokens
    if token.pos in {'ADJ', 'ADJF', 'ADJS'}
    and token.lemma not in custom_exclude
]

verb_freq = FreqDist(verbs)
adj_freq = FreqDist(adjectives)

print("\n🔧 Топ-10 глаголов:")
for word, freq in verb_freq.most_common(10):
    print((word, freq))

print("\n🎨 Топ-10 прилагательных:")
for word, freq in adj_freq.most_common(10):
    print((word, freq))

print("Найдено прилагательных:", len(adjectives))
print(adjectives[:10])

# === 8. Построение облаков слов ===
font_path = "C:/Windows/Fonts/arial.ttf"  # ⚠️ укажи подходящий шрифт с кириллицей

wordcloud_all = WordCloud(
    width=1000, height=700,
    background_color="white",
    colormap="viridis",
    font_path=font_path,
    max_words=100
).generate_from_frequencies(fdist)

wordcloud_verbs = WordCloud(
    width=1000, height=700,
    background_color="white",
    colormap="plasma",
    font_path=font_path,
    max_words=100
).generate_from_frequencies(verb_freq)

wordcloud_adjs = WordCloud(
    width=1000, height=700,
    background_color="white",
    colormap="inferno",
    font_path=font_path,
    max_words=100
).generate_from_frequencies(adj_freq)

# === 9. Визуализация и сохранение ===
plt.figure(figsize=(18, 12))

plt.subplot(1, 3, 1)
plt.imshow(wordcloud_all, interpolation="bilinear")
plt.axis("off")
plt.title("Все слова", fontsize=16)

plt.subplot(1, 3, 2)
plt.imshow(wordcloud_verbs, interpolation="bilinear")
plt.axis("off")
plt.title("Глаголы", fontsize=16)

plt.subplot(1, 3, 3)
plt.imshow(wordcloud_adjs, interpolation="bilinear")
plt.axis("off")
plt.title("Прилагательные", fontsize=16)

plt.tight_layout()
plt.show()
