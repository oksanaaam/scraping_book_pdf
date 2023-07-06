import openpyxl
import fitz


pdf_file = "C:/Users/User/projects/scraping_book_pdf/Abstract Book.pdf"
xl_file = "C:/Users/User/projects/scraping_book_pdf/task.xlsx"

start_page = 44
last_page = 64


class Article:
    def __init__(self):
        self.session_name = ""
        self.session_title = ""
        self.authors = []
        self.affiliations = []
        self.presentation_abstract = ""
        self.location = []

    def __str__(self):
        return (
            f"\nSession name: {self.session_name}"
            f"\nSession title: {self.session_title}"
            f"\nAuthors: {', '.join(self.authors)}"
            f"\nAffiliations: {', '.join(self.affiliations)}"
            f"\nLocation: {', '.join(self.location)}"
            f"\nPresentation abstract: {self.presentation_abstract}"
        )


def extract_articles_from_pdf(file_path: str):
    articles = []
    current_article = Article()
    previous_article = None

    file = fitz.open(file_path)

    for page_number in range(start_page - 1, last_page):
        page = file[page_number]
        text_blocks = page.get_text("dict")["blocks"]

        for block in text_blocks:
            for line in block["lines"]:
                for span in line["spans"]:
                    block_name = (
                        span["font"] == "TimesNewRomanPS-BoldItal"
                        and span["size"] == 9.5
                    )
                    block_title = (
                        span["font"] == "TimesNewRomanPS-BoldMT"
                        and span["size"] == 9
                    )
                    authors = (
                        span["font"] == "TimesNewRomanPS-ItalicMT"
                        and span["size"] == 9
                    )
                    affiliations_data = (
                        span["font"] == "TimesNewRomanPS-ItalicMT"
                        and span["size"] == 8
                    )
                    abstract_data = span["size"] == 9.134002685546875

                    text_pattern = span["text"]

                    if block_name:
                        current_article.session_name += text_pattern
                    elif block_title:
                        current_article.session_title += text_pattern
                    elif authors:
                        author = text_pattern.strip().replace(", ", "")
                        print(author)
                        if author:
                            current_article.authors.append(author)
                    elif affiliations_data:
                        if "," in text_pattern:
                            words = text_pattern.split(",")
                            affiliations = []
                            location = ""
                            for word in words:
                                word = word.strip()
                                if word:
                                    if len(word.split()) == 1:
                                        location += word + ", "
                                    else:
                                        affiliations.append(word)
                            location = location.rstrip(", ")
                            if location:
                                current_article.location.append(location)
                            current_article.affiliations.extend(affiliations)
                        else:
                            if text_pattern.strip():
                                current_article.affiliations.append(text_pattern)

                    elif abstract_data:
                        current_article.presentation_abstract += text_pattern

            if block["type"] == 0:
                if (
                    current_article.session_name
                    or current_article.session_title
                    or current_article.authors
                    or current_article.affiliations
                    or current_article.presentation_abstract
                ):
                    if not current_article.session_name:
                        if previous_article is not None:
                            previous_article.session_title += current_article.session_title
                            previous_article.affiliations.extend(current_article.affiliations)
                            previous_article.location.extend(current_article.location)
                            previous_article.presentation_abstract += current_article.presentation_abstract
                            current_article = previous_article
                    articles.append(current_article)
                    previous_article = current_article
                    current_article = Article()

    file.close()
    return articles


def update_excel_file(article):
    workbook = openpyxl.load_workbook(xl_file)
    sheet = workbook.active

    row = sheet.max_row + 1

    sheet.cell(row=row, column=1).value = ", ".join(article.authors)
    sheet.cell(row=row, column=2).value = ", ".join(article.affiliations)
    sheet.cell(row=row, column=3).value = ", ".join(article.location)
    sheet.cell(row=row, column=4).value = article.session_name
    sheet.cell(row=row, column=5).value = article.session_title
    sheet.cell(row=row, column=6).value = article.presentation_abstract

    workbook.save(xl_file)


if __name__ == "__main__":
    pdf_articles = extract_articles_from_pdf(pdf_file)
    for article in pdf_articles:
        print(article)
        update_excel_file(article)
