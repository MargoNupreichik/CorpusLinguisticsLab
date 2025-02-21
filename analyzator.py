# Импортируем необходимые библиотеки
from textblob import TextBlob
import pandas as pd


def clean_review(review):
    """
    Функция для очистки отзыва от ненужных символов.
    
    :param review: Исходный отзыв в виде строки
    :return: Очищенный отзыв
    """
    # Удаляем все символы, кроме букв и пробелов
    cleaned_review = ''.join(char for char in review if char.isalnum() or char.isspace())
    return cleaned_review

def analyze_sentiment(review):
    """
    Функция для анализа тональности отзыва.
    
    :param review: Очищенный отзыв в виде строки
    :return: Объект TextBlob с анализом тональности
    """
    # Создаём объект TextBlob из очищенного отзыва
    blob = TextBlob(review)
    
    polarity = blob.sentiment.polarity
    polarity_grade = 'positive' if polarity > 0.25 else 'negative' if polarity < -0.25 else 'neutral'
    
    subjectivity = blob.sentiment.subjectivity
    subjectivity_grade = 'subjective' if subjectivity > 0.5 else 'objective'
    
    return pd.Series({'polarity': polarity,
                      'polarity_grade': polarity_grade, 
                      'subjectivity': subjectivity,
                      'subjectivity_grade': subjectivity_grade})
     

# Пример использования функций
if __name__ == "__main__":
    # Фрейм с исходными отзывами
    df = pd.read_excel('reviews.xlsx')
    
    # Очищаем отзывы от ненужных символов
    df['cleaned_review'] = df['review'].apply(clean_review)
    
    # Анализируем тональность очищенных отзывов
    df[['polarity', 'polarity_grade', 'subjectivity', 'subjectivity_grade']] = df['cleaned_review'].apply(analyze_sentiment)
    
    # Сохраняем результат в новый файл Excel
    with pd.ExcelWriter('reviews_with_sentiment.xlsx') as writer:
        df.to_excel(writer, sheet_name='All')
        df[df['polarity_grade'] == 'positive'].to_excel(writer, sheet_name='Positive', index=False)
        df[df['polarity_grade'] == 'neutral'].to_excel(writer, sheet_name='Neutral', index=False)
        df[df['polarity_grade'] == 'negative'].to_excel(writer, sheet_name='Negative', index=False)
        df[df['subjectivity_grade'] == 'subjective'].to_excel(writer, sheet_name='Subjective', index=False)
        df[df['subjectivity_grade'] == 'objective'].to_excel(writer, sheet_name='Objective', index=False)
    
    print('Анализ завершен. Результаты сохранены в файл "reviews_with_sentiment.xlsx".')