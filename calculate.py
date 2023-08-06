import openpyxl
import itertools
import jaconv
from util import get_list_2d

# 参加者の所属チーム
teams = [
  ["Ryosuke Hyakuta", "Yuta Sakai", "Hiroki Takeuchi"],
  ["TestA", "TestB", "TestC"],
]

# 参加者の回答を格納する
participants = {
  "TestA": [],
  "TestC": [],
  "TestB": [],
  "Yuta Sakai": [],
  "Ryosuke Hyakuta": [],
  "Hiroki Takeuchi": [],
}

# チーム別の回答
team_answers = []

# チーム別の得点
team_scores = []

# スペシャルワード
special_words = [
  "まいくろそふと",  # マイクロソフト
  "あじゅーる",  # Azure
  "わーど",  # Word
  "ぱわーぽいんと",  # PowerPoint
  "うぃんどうず"  # Windows
  "ひっとりふれっしゅ"  # Hit Refresh
]

# スペシャルワードによる得点
SPECIAL_WORD_POINTS = 10

FIRST_WORD = "くらうど"

def load_sheet(file_path: str) -> openpyxl.worksheet.worksheet.Worksheet:
    wb = openpyxl.load_workbook(file_path)
    return wb["Sheet1"]


def collect_answers(sheet: openpyxl.worksheet.worksheet.Worksheet) -> None:
    i = 2
    while (sheet["G{}".format(i)].value is not None):

        name = sheet["G{}".format(i)].value
        i += 1

        answers = get_list_2d(sheet, i-1, i-1, 8, 22)[0]
        answers = [item for item in answers if item is not None]
        participants[name] = answers


def collect_team_answers() -> None:
    for team in teams:
        _team_answers = []
        for name in team:
            _team_answers.append(participants[name])

        zipped = list(itertools.zip_longest(*_team_answers))
        new_zipped = [x for t in zipped for x in t if x is not None]
        team_answers.append(new_zipped)


def calculate_score(answers: list) -> int:
    # print(answers)

    # カタカナ→ひらがなに変換
    hiragana_answers = [jaconv.kata2hira(item) for item in answers]
    # print(hiragana_answers)

    # 長音を母音に変換
    vowelized_answers = replace_with_vowel(hiragana_answers)
    # print(vowelized_answers)

    # 拗音・促音（ぁ、ぃ、ぅ、ぇ、ぉ、ゃ、ゅ、ょ、っ）を清音に変換
    cleared_answers = replace_yoon_sokuon(vowelized_answers)
    # 最初のお題を追加
    cleared_answers.insert(0, FIRST_WORD)
    # print(cleared_answers)

    # 連鎖率による得点を計算
    (
        score_by_consecutive,
        special_word_count,
        correct_word_len
    ) = CheckShiritori(cleared_answers)
    # print(score_by_consecutive)

    # スペシャルワードによる得点を計算
    special_word_score = special_word_count * SPECIAL_WORD_POINTS
    # print(special_word_score)

    # 正答文字数による得点を計算
    # print(correct_word_len)

    # 合計得点を計算
    # print(score_by_consecutive,
    #     special_word_score,
    #     correct_word_len)
    total_score = score_by_consecutive + special_word_score + correct_word_len
    return total_score


def convert_to_vowel(hiragana: str) -> str:
    vowel_map = {
        "あ": "あ", "い": "い", "う": "う", "え": "え", "お": "お",
        "か": "あ", "き": "い", "く": "う", "け": "え", "こ": "お",
        "さ": "あ", "し": "い", "す": "う", "せ": "え", "そ": "お",
        "た": "あ", "ち": "い", "つ": "う", "て": "え", "と": "お",
        "な": "あ", "に": "い", "ぬ": "う", "ね": "え", "の": "お",
        "は": "あ", "ひ": "い", "ふ": "う", "へ": "え", "ほ": "お",
        "ま": "あ", "み": "い", "む": "う", "め": "え", "も": "お",
        "や": "あ", "ゆ": "う", "よ": "お",
        "ら": "あ", "り": "い", "る": "う", "れ": "え", "ろ": "お",
        "わ": "あ", "を": "お", "ん": "ん",
        "が": "あ", "ぎ": "い", "ぐ": "う", "げ": "え", "ご": "お",
        "ざ": "あ", "じ": "い", "ず": "う", "ぜ": "え", "ぞ": "お",
        "だ": "あ", "ぢ": "い", "づ": "う", "で": "え", "ど": "お",
        "ば": "あ", "び": "い", "ぶ": "う", "べ": "え", "ぼ": "お",
        "ぱ": "あ", "ぴ": "い", "ぷ": "う", "ぺ": "え", "ぽ": "お",
        "ゃ": "あ", "ゅ": "う", "ょ": "お"
    }

    return vowel_map[hiragana] if hiragana in vowel_map else hiragana


def replace_with_vowel(words: list) -> list:
    for i, word in enumerate(words):
        if word.endswith("ー"):
            # 長音の前の文字を取得し、それを母音に変換
            vowel = convert_to_vowel(word[-2])
            # 長音を母音に置換
            words[i] = word[:-1] + vowel
    return words


def replace_yoon_sokuon(words: list) -> list:
    yoon_sokuon_to_seion = {
        "ぁ": "あ", "ぃ": "い", "ぅ": "う", "ぇ": "え", "ぉ": "お",
        "ゃ": "や", "ゅ": "ゆ", "ょ": "よ", "っ": "つ"
    }
    for i, word in enumerate(words):
        if word[-1] in yoon_sokuon_to_seion:
            # 拗音・促音の文字を取得し、それを清音に変換
            seion = yoon_sokuon_to_seion[word[-1]]
            # 拗音・促音を清音に置換
            words[i] = word[:-1] + seion
    return words


def CheckShiritori(words: list) -> tuple[int, int, int]:
    score = 0
    consecutive_correct = 0
    used_words = set()
    special_word_count = 0
    correct_word_len = 0

    for i in range(len(words) - 1):
        # 同じ単語が再利用されていないかチェック
        if words[i] in used_words:
            consecutive_correct = 0
        else:
            used_words.add(words[i])
            # 前の単語の終わりの文字と次の単語の最初の文字が一致しているかチェック
            if words[i][-1] == words[i+1][0]:
                consecutive_correct += 1
                score += consecutive_correct * (consecutive_correct + 1) // 2
                # スペシャルワードの判定
                if words[i+1] in special_words:
                    special_word_count += 1

                correct_word_len += len(words[i+1])
            else:
                consecutive_correct = 0

    return score, special_word_count, correct_word_len


if __name__ == "__main__":

    # エクセルファイルを読み込む
    sheet = load_sheet("data/results.xlsx")

    # 回答を集計する
    collect_answers(sheet)

    # チーム別の回答を集計する
    collect_team_answers()

    # チーム別の得点を計算する
    for answers in team_answers:
        team_scores.append(calculate_score(answers))

    # 結果を表示する
    for i in range(len(teams)):
        print("-----------------------------------------------------------")
        print("チーム{} : {}".format(i+1, teams[i]))
        print("得点 : {}".format(team_scores[i]))
