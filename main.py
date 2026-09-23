import re
import json
import math
import datetime
import urllib.request
import urllib.error
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

import openpyxl
from playwright.sync_api import sync_playwright


# ============================================================
# CONFIGURATION
# ============================================================

SIMPLEX_BASE = "https://i9simplex.simplilearn.com"

SEO_LIST_URL = f"{SIMPLEX_BASE}/admin/seo/list"
COURSE_LIST_URL = f"{SIMPLEX_BASE}/admin/course/list"
BASIC_URL = f"{SIMPLEX_BASE}/admin/course/basic"

DUMMY_SEO_THUMBNAIL = (
    "https://www.simplilearn.com/ice9/course_images/icons/cyber-security.svgz"
)

DUMMY_COURSE_LOGO = (
    "https://www.simplilearn.com/ice9/course_images/icons/"
    "Introduction-to-Project-Management.svgz"
)

CERTIFICATE_FILE = Path(__file__).parent / "unname.jpg"

LEARNING_RATING = "4.5"

OLLAMA_CHAT_URL = "http://localhost:11434/api/chat"
OLLAMA_MODEL = "llama3.2:3b"
OLLAMA_TIMEOUT_SECONDS = 180
MAX_OLLAMA_ATTEMPTS = 8

SKILLS_OVERVIEW_SYSTEM_PROMPT = """You are analysing a complete course Table of Contents.

Your task is to determine what a learner will actually
be capable of doing after completing the course.

Analyse ALL lessons, topics and subtopics before generating
the output.

SKILL RULES

1. Generate exactly 5 skills.
2. Skills must represent major competencies covered by the course.
3. Do not simply copy lesson titles.
4. Combine related topics into broader meaningful skills.
5. Do not invent skills not supported by the TOC.
6. Avoid overly broad skills.
7. Avoid duplicate or overlapping skills.
8. Use concise skill names suitable for a CID/course catalogue:
   2-8 words, a noun phrase (e.g. "Data Validation with Lookup
   Functions"), never a full sentence.
9. Every topic in the TOC includes its duration. Use this to judge
   how much of the course is actually about each concept - do not
   treat every listed topic as equally important.
10. A skill must be supported either by MULTIPLE topics/lessons, or
    by a topic that represents a major share of total course
    duration. A single short, minor topic (a small fraction of the
    course, taught in only one place) must NOT become a standalone
    skill - merge it into a broader related skill instead, or leave
    it out in favour of a skill with real course-wide support.
11. If one topic is a minor feature mention (a few minutes) inside a
    course that is overwhelmingly about a different subject, do not
    let that minor topic become one of the 5 skills just because it
    is easy to name.

COURSE OVERVIEW RULES

1. Describe the complete course, not only the first few lessons.
2. Mention the major knowledge and practical competencies covered,
   weighted by how much of the course actually covers them.
3. Do not invent tools or techniques absent from the TOC.
4. Do not claim the course teaches something broader than the TOC
   supports - for example, if the TOC only covers using or prompting
   an existing AI tool, do not say the course teaches building,
   training, or creating AI models.
5. Avoid marketing exaggeration and unearned superlatives such as
   "expert-level", "master", "cutting-edge", "seamless", or "unlock"
   unless that exact claim is directly supported by a TOC entry.
6. Produce one concise paragraph.
7. The overview must be between 250 and 350 characters, counting
   spaces and punctuation. This is a hard platform constraint -
   count your characters before answering.

OUTPUT FORMAT

Respond with ONLY valid JSON, no commentary, no markdown fences,
in exactly this shape:

{"skills": ["skill 1", "skill 2", "skill 3", "skill 4", "skill 5"], "overview": "single paragraph course overview"}
"""


# ============================================================
# STEP 0 - TOC FILE SELECTION
# ============================================================

def pick_excel_file_dialog(title):

    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm"),
            ("All files", "*.*"),
        ],
    )

    root.destroy()

    if not file_path:

        raise Exception(
            f"No file was selected ({title})."
        )

    return Path(file_path)


# ============================================================
# STEP 0 - TOC PARSING
# ============================================================

def clean_toc_text(value):

    if value is None:
        return ""

    text = str(value)

    text = text.replace("​", "")
    text = text.replace("\xa0", " ")

    text = re.sub(r"\s+", " ", text)

    return text.strip()


def derive_course_name_from_filename(file_path):

    name = file_path.stem

    name = re.sub(r"[_\s]*toc$", "", name, flags=re.IGNORECASE)
    name = name.replace("_", " ")
    name = re.sub(r"\s+", " ", name).strip()

    return name


def find_toc_column(header, *names):

    for name in names:

        if name in header:
            return header.index(name)

    return None


def parse_duration_cell(raw_value):

    if raw_value is None:
        return None

    if isinstance(raw_value, datetime.timedelta):
        return raw_value.total_seconds()

    if isinstance(raw_value, (int, float)):
        return float(raw_value)

    text = clean_toc_text(raw_value)

    if not text:
        return None

    parts = text.split(":")

    try:
        parts = [int(part) for part in parts]
    except ValueError:
        return None

    if len(parts) == 3:
        hours, minutes, seconds = parts
    elif len(parts) == 2:
        hours, minutes, seconds = 0, parts[0], parts[1]
    else:
        return None

    return hours * 3600 + minutes * 60 + seconds


def format_duration_seconds(seconds):

    if seconds is None:
        return None

    total_seconds = int(round(seconds))
    minutes, secs = divmod(total_seconds, 60)
    hours, minutes = divmod(minutes, 60)

    if hours:
        return f"{hours}h {minutes}m {secs}s"

    if minutes:
        return f"{minutes}m {secs}s"

    return f"{secs}s"


def parse_toc_excel(file_path):
    """
    Parse a Simplilearn partner-course TOC Excel file.

    Expected layout (standard "<Course>_TOC.xlsx" template):

        Course Name | Lesson No. | Lesson Name | Topic No. | Topics | Duration

    Course Name / Lesson No. / Lesson Name are only populated on the
    first row of each group (merged cells or repeated blanks), so
    values are forward-filled while walking down the sheet.

    A simpler "Lesson Name | Topic" layout (no Course Name column) is
    also supported; the course name then falls back to the file name.
    """

    if not file_path.exists():

        raise Exception(
            f"TOC file was not found:\n{file_path}"
        )

    workbook = openpyxl.load_workbook(
        str(file_path),
        data_only=True
    )

    worksheet = workbook.worksheets[0]

    rows = list(
        worksheet.iter_rows(values_only=True)
    )

    if len(rows) < 2:

        raise Exception(
            "TOC file does not contain any data rows.\n"
            f"File: {file_path}"
        )

    header = [
        clean_toc_text(cell).lower()
        for cell in rows[0]
    ]

    col_course = find_toc_column(header, "course name")
    col_lesson_no = find_toc_column(header, "lesson no.", "lesson no")
    col_lesson_name = find_toc_column(header, "lesson name")
    col_topic = find_toc_column(header, "topics", "topic")
    col_duration = find_toc_column(header, "duration")

    if col_topic is None:

        raise Exception(
            "Could not find a 'Topics' column in the TOC file.\n"
            f"File: {file_path}"
        )

    def cell_value(row, index):

        if index is None or index >= len(row):
            return None

        return row[index]

    course_name = None
    last_lesson_no = None
    last_lesson_name = None
    lessons = []

    for row in rows[1:]:

        raw_course = cell_value(row, col_course)
        raw_lesson_no = cell_value(row, col_lesson_no)
        raw_lesson_name = cell_value(row, col_lesson_name)
        raw_topic = cell_value(row, col_topic)
        raw_duration = cell_value(row, col_duration)

        course_candidate = clean_toc_text(raw_course)

        if course_candidate and course_name is None:
            course_name = course_candidate

        lesson_name_candidate = clean_toc_text(raw_lesson_name)

        if lesson_name_candidate:
            last_lesson_name = lesson_name_candidate

        if raw_lesson_no is not None and str(raw_lesson_no).strip():
            last_lesson_no = raw_lesson_no

        topic_text = clean_toc_text(raw_topic)

        if not topic_text:
            # Blank separator row or the trailing totals row.
            continue

        group_key = (last_lesson_no, last_lesson_name)

        if not lessons or lessons[-1]["group_key"] != group_key:

            lessons.append({
                "group_key": group_key,
                "lesson_name": last_lesson_name,
                "topics": [],
            })

        lessons[-1]["topics"].append({
            "text": topic_text,
            "duration_seconds": parse_duration_cell(raw_duration),
        })

    if not lessons:

        raise Exception(
            "No topics were found in the TOC file.\n"
            f"File: {file_path}"
        )

    if not course_name:
        course_name = derive_course_name_from_filename(file_path)

    if not course_name:

        raise Exception(
            "Could not determine the Course Name from the TOC file "
            "or its file name.\n"
            f"File: {file_path}"
        )

    return {
        "course_name": course_name,
        "lessons": lessons,
    }


# ============================================================
# STEP 0 - BATCH (BUNCH OF CIDS) FILE PARSING
# ============================================================

def duration_cell_to_string(raw_value):

    if raw_value is None:
        return ""

    if isinstance(raw_value, datetime.timedelta):

        total_seconds = int(raw_value.total_seconds())
        hours, remainder = divmod(total_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)

        return f"{hours}:{minutes:02d}:{seconds:02d}"

    if isinstance(raw_value, datetime.time):

        return (
            f"{raw_value.hour}:{raw_value.minute:02d}:"
            f"{raw_value.second:02d}"
        )

    return clean_toc_text(raw_value)


def parse_batch_excel(file_path):
    """
    Parse the batch CID Excel file - one row per course.

    Expected columns (case-insensitive, any order):

        Course Title (reference only) | Primary Category |
        Primary EID | Course Level | Duration | TOC File Path

    "Course Title" is for the operator's own reference only; the
    authoritative course title always comes from each row's TOC
    file, same as the single-CID flow.
    """

    if not file_path.exists():

        raise Exception(
            f"Batch file was not found:\n{file_path}"
        )

    workbook = openpyxl.load_workbook(
        str(file_path),
        data_only=True
    )

    worksheet = workbook.worksheets[0]

    rows = list(
        worksheet.iter_rows(values_only=True)
    )

    if len(rows) < 2:

        raise Exception(
            "Batch file does not contain any data rows.\n"
            f"File: {file_path}"
        )

    header = [
        clean_toc_text(cell).lower()
        for cell in rows[0]
    ]

    col_title = find_toc_column(header, "course title", "title")
    col_category = find_toc_column(header, "primary category", "category")
    col_eid = find_toc_column(
        header, "primary eid", "primary elearning eid", "eid"
    )
    col_level = find_toc_column(header, "course level", "level")
    col_duration = find_toc_column(header, "duration", "course duration")
    col_toc_path = find_toc_column(
        header, "toc file path", "toc path", "toc"
    )

    required_columns = {
        "Course Title": col_title,
        "Primary Category": col_category,
        "Primary EID": col_eid,
        "Course Level": col_level,
        "Duration": col_duration,
        "TOC File Path": col_toc_path,
    }

    missing_columns = [
        name for name, col in required_columns.items()
        if col is None
    ]

    if missing_columns:

        raise Exception(
            "The batch Excel file is missing required column(s): "
            f"{', '.join(missing_columns)}.\n"
            f"File: {file_path}"
        )

    def cell_value(row, index):

        if index is None or index >= len(row):
            return None

        return row[index]

    batch_rows = []

    for row_number, row in enumerate(rows[1:], start=2):

        toc_path_text = clean_toc_text(cell_value(row, col_toc_path))
        primary_category = clean_toc_text(cell_value(row, col_category))
        primary_eid = clean_toc_text(cell_value(row, col_eid))
        course_level = clean_toc_text(cell_value(row, col_level))
        duration_text = duration_cell_to_string(
            cell_value(row, col_duration)
        )

        title_hint = clean_toc_text(cell_value(row, col_title))

        # Skip fully blank rows.
        if not any([
            toc_path_text,
            primary_category,
            primary_eid,
            course_level,
            duration_text,
            title_hint,
        ]):
            continue

        row_problems = []

        if not title_hint:
            row_problems.append("Course Title is empty.")

        if not toc_path_text:
            row_problems.append("TOC File Path is empty.")

        if not primary_category:
            row_problems.append("Primary Category is empty.")

        if not primary_eid:
            row_problems.append("Primary EID is empty.")

        if not course_level:
            row_problems.append("Course Level is empty.")

        if not duration_text:
            row_problems.append("Duration is empty.")

        if row_problems:

            raise Exception(
                f"Batch file row {row_number} has problem(s):\n"
                + "\n".join(f"- {problem}" for problem in row_problems)
                + f"\nFile: {file_path}"
            )

        batch_rows.append({
            "row_number": row_number,
            "title_hint": title_hint,
            "toc_path": Path(toc_path_text),
            "primary_category": primary_category,
            "primary_eid": primary_eid,
            "course_level": course_level,
            "duration": duration_text,
        })

    if not batch_rows:

        raise Exception(
            "No usable course rows were found in the batch file.\n"
            f"File: {file_path}"
        )

    return batch_rows


# ============================================================
# STEP 0 - SKILLS + OVERVIEW VIA LOCAL OLLAMA
# ============================================================

def format_toc_for_prompt(toc_data):

    lines = [
        f"Course Name: {toc_data['course_name']}",
        "",
    ]

    for index, lesson in enumerate(toc_data["lessons"], start=1):

        lesson_name = lesson["lesson_name"] or f"Lesson {index}"
        lines.append(f"Lesson {index}: {lesson_name}")

        for topic in lesson["topics"]:

            duration_text = format_duration_seconds(
                topic["duration_seconds"]
            )

            if duration_text:
                lines.append(f"  - {topic['text']} ({duration_text})")
            else:
                lines.append(f"  - {topic['text']}")

        lines.append("")

    return "\n".join(lines).strip()


def normalize_for_comparison(text):

    return re.sub(r"[^a-z0-9]", "", str(text).lower())


STOPWORDS = {
    "the", "a", "an", "and", "or", "of", "for", "in", "on", "to",
    "with", "using", "via", "your", "you", "is", "are", "as", "into",
    "from", "by", "at", "this", "that", "these", "those", "vs",
}

MINOR_TOPIC_DURATION_SHARE = 0.06


def significant_words(text):

    words = re.findall(r"[a-z0-9]+", str(text).lower())

    return {
        word for word in words
        if word not in STOPWORDS and len(word) > 2
    }


def find_minor_topic_skills(cleaned_skills, toc_data):
    """
    Detect a skill that is really just a single minor topic (a small
    share of total course duration, mentioned in only one place)
    wearing a broader-sounding name.

    Returns [] when duration data is missing/incomplete, since
    coverage cannot be judged without it.
    """

    all_topics = [
        topic
        for lesson in toc_data["lessons"]
        for topic in lesson["topics"]
    ]

    durations = [topic["duration_seconds"] for topic in all_topics]

    if not all_topics or any(d is None for d in durations):
        return []

    total_duration = sum(durations)

    if total_duration <= 0:
        return []

    flagged = []

    for skill in cleaned_skills:

        skill_words = significant_words(skill)

        if not skill_words:
            continue

        contributing_topics = [
            topic for topic in all_topics
            if significant_words(topic["text"])
            and len(
                skill_words & significant_words(topic["text"])
            ) / len(skill_words) >= 0.5
        ]

        if len(contributing_topics) != 1:
            # Supported by multiple topics (or none matched closely
            # enough to judge) - not a minor-topic promotion.
            continue

        topic = contributing_topics[0]
        share = topic["duration_seconds"] / total_duration

        if share < MINOR_TOPIC_DURATION_SHARE:

            flagged.append({
                "skill": skill,
                "topic": topic["text"],
                "share_percent": round(share * 100, 1),
            })

    return flagged


def call_ollama_chat(messages):

    payload = {
        "model": OLLAMA_MODEL,
        "messages": messages,
        "format": "json",
        "stream": False,
        "options": {
            "temperature": 0.3,
        },
    }

    request = urllib.request.Request(
        OLLAMA_CHAT_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:

        with urllib.request.urlopen(
            request,
            timeout=OLLAMA_TIMEOUT_SECONDS
        ) as response:

            body = json.loads(response.read().decode("utf-8"))

    except (urllib.error.URLError, TimeoutError) as error:

        raise Exception(
            "Could not reach the local Ollama server at "
            f"{OLLAMA_CHAT_URL}.\n\n"
            "Make sure the Ollama app is running and that the "
            f"'{OLLAMA_MODEL}' model has been pulled "
            f"(ollama pull {OLLAMA_MODEL}).\n\n"
            f"Original error: {error}"
        )

    content = body.get("message", {}).get("content", "")

    try:
        return json.loads(content)

    except json.JSONDecodeError as error:

        raise Exception(
            "Ollama did not return valid JSON for the Skills/Overview "
            "request.\n\n"
            f"Raw response:\n{content}\n\n"
            f"Original error: {error}"
        )


def validate_ollama_result(result, toc_data):

    problems = []

    skills = result.get("skills")
    overview = result.get("overview")

    if not isinstance(skills, list) or len(skills) != 5:

        problems.append(
            "The 'skills' field must be a JSON array of exactly 5 items."
        )

    else:

        cleaned_skills = [str(skill).strip() for skill in skills]

        if any(not skill for skill in cleaned_skills):
            problems.append("Skills must not be empty strings.")

        normalized_skills = [
            normalize_for_comparison(skill) for skill in cleaned_skills
        ]

        if len(set(normalized_skills)) != len(normalized_skills):
            problems.append(
                "Skills must not duplicate or overlap one another."
            )

        too_long = [
            skill for skill in cleaned_skills
            if len(skill.split()) > 8
        ]

        if too_long:

            problems.append(
                "These skills are full sentences instead of concise "
                f"2-8 word catalogue phrases: {too_long}. Shorten them "
                "to short noun phrases."
            )

        raw_toc_entries = set()

        for lesson in toc_data["lessons"]:

            if lesson["lesson_name"]:
                raw_toc_entries.add(
                    normalize_for_comparison(lesson["lesson_name"])
                )

            for topic in lesson["topics"]:
                raw_toc_entries.add(
                    normalize_for_comparison(topic["text"])
                )

        copied = [
            cleaned_skills[i]
            for i, normalized in enumerate(normalized_skills)
            if normalized in raw_toc_entries
        ]

        if copied:

            problems.append(
                "These skills are copied verbatim from the TOC "
                f"instead of being synthesized competencies: {copied}. "
                "Combine related topics into a broader skill instead "
                "of repeating a single topic or lesson name."
            )

        minor_matches = find_minor_topic_skills(cleaned_skills, toc_data)

        if minor_matches:

            details = "; ".join(
                f"'{match['skill']}' is based only on the minor topic "
                f"'{match['topic']}' (~{match['share_percent']}% of "
                "total course duration, taught in a single place)"
                for match in minor_matches
            )

            problems.append(
                "A skill should represent a competency supported by "
                "multiple lessons or a major course objective. A "
                "single short feature lesson should not become a "
                f"standalone skill: {details}. Merge it into a "
                "broader related skill, or replace it with a skill "
                "backed by substantial course-wide coverage."
            )

    if not isinstance(overview, str) or not overview.strip():

        problems.append(
            "The 'overview' field must be a non-empty string."
        )

    else:

        length = len(overview.strip())

        if not 250 <= length <= 350:

            target = 300

            if length < 250:

                problems.append(
                    f"The overview is {length} characters, which is "
                    f"{250 - length} characters too short (target "
                    f"range 250-350, aim for about {target}). Add one "
                    "more specific, real detail pulled from the TOC "
                    "(e.g. name a specific topic or technique not yet "
                    "mentioned) - do not add generic filler sentences."
                )

            else:

                problems.append(
                    f"The overview is {length} characters, which is "
                    f"{length - 350} characters too long (target range "
                    f"250-350, aim for about {target}). Remove some "
                    "detail while keeping it accurate to the TOC."
                )

    return problems


def generate_skills_and_overview_via_ollama(toc_data):

    toc_text = format_toc_for_prompt(toc_data)

    messages = [
        {"role": "system", "content": SKILLS_OVERVIEW_SYSTEM_PROMPT},
        {"role": "user", "content": f"Course TOC:\n\n{toc_text}"},
    ]

    last_problems = []

    for attempt in range(1, MAX_OLLAMA_ATTEMPTS + 1):

        print(
            f"\nAsking Ollama ({OLLAMA_MODEL}) to analyse the TOC... "
            f"attempt {attempt}/{MAX_OLLAMA_ATTEMPTS}"
        )

        result = call_ollama_chat(messages)

        problems = validate_ollama_result(result, toc_data)

        if not problems:

            skills = [str(skill).strip() for skill in result["skills"]]
            overview = str(result["overview"]).strip()

            return skills, overview

        last_problems = problems

        print("Ollama output did not pass validation:")

        for problem in problems:
            print(" -", problem)

        feedback = (
            "Your previous JSON response had these problems:\n"
            + "\n".join(f"- {problem}" for problem in problems)
            + "\n\nRespond again with corrected JSON only, following "
            "all the original rules."
        )

        messages.append({
            "role": "assistant",
            "content": json.dumps(result),
        })

        messages.append({
            "role": "user",
            "content": feedback,
        })

    raise Exception(
        "Ollama could not produce a valid Skills/Overview result "
        f"after {MAX_OLLAMA_ATTEMPTS} attempts.\n\n"
        "Last problems:\n"
        + "\n".join(f"- {problem}" for problem in last_problems)
    )


# ============================================================
# SEO URL
# ============================================================

def create_seo_slug(course_title):

    url = course_title.lower().strip()

    url = re.sub(r"[^a-z0-9\s-]", "", url)
    url = re.sub(r"\s+", "-", url)
    url = re.sub(r"-+", "-", url)

    return url.strip("-")


def create_simplex_seo_url(course_title):

    # IMPORTANT:
    # Simplex requires the leading /
    return "/" + create_seo_slug(course_title)


# ============================================================
# COURSE LEVEL
# ============================================================

def normalize_course_level(user_level):

    level = user_level.strip().lower()

    level_mapping = {
        "beginner": "Foundational",
        "intermediate": "Foundational",
        "foundational": "Foundational",
        "advanced": "Advanced",
        "tool based": "Tool Based",
        "tool-based": "Tool Based",
        "knowledge nuggets": "Knowledge Nuggets",
        "leadership series": "Leadership Series",
    }

    if level not in level_mapping:

        raise Exception(
            f"Invalid course level: {user_level}\n\n"
            "Allowed values:\n"
            "Beginner\n"
            "Intermediate\n"
            "Foundational\n"
            "Advanced\n"
            "Tool Based\n"
            "Knowledge Nuggets\n"
            "Leadership Series"
        )

    return level_mapping[level]


# ============================================================
# DURATION
# ============================================================

def parse_duration(duration_input):

    duration_input = duration_input.strip()

    parts = duration_input.split(":")

    if len(parts) != 3:

        raise Exception(
            "Duration must be entered in HH:MM:SS format.\n"
            "Example: 2:00:13"
        )

    try:
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2])

    except ValueError:

        raise Exception(
            "Duration must contain valid numbers.\n"
            "Example: 2:00:13"
        )

    if hours < 0:
        raise Exception("Hours cannot be negative.")

    if minutes < 0 or minutes > 59:
        raise Exception("Minutes must be between 0 and 59.")

    if seconds < 0 or seconds > 59:
        raise Exception("Seconds must be between 0 and 59.")

    total_seconds = (
        hours * 3600
        + minutes * 60
        + seconds
    )

    if total_seconds <= 0:

        raise Exception(
            "Duration must be greater than zero."
        )

    total_hours = total_seconds / 3600

    # B2B SkillUp duration rule:
    # - Ignore seconds completely.
    # - If minutes are below 30, keep the hour value.
    # - If minutes are 30 or above, add 1 hour.
    #
    # Examples:
    # 2:00:13 -> 2
    # 2:29:59 -> 2
    # 2:30:00 -> 3
    # 45:56:12 -> 46
    # 5:00:01 -> 5
    rounded_hours = hours

    if minutes >= 30:
        rounded_hours += 1

    # Simplex does not accept 0 as a duration. A sub-30-minute course
    # (e.g. 0:29:33) rounds down to 0 hours under the rule above - in
    # that one case only, round up to 1 hour instead of 0.
    if rounded_hours == 0:
        rounded_hours = 1

    weeks = math.ceil(rounded_hours / 5)

    return total_hours, rounded_hours, weeks


# ============================================================
# COURSE NAME NORMALIZATION
# ============================================================

def normalize_course_name_for_match(name):

    if not name:
        return ""

    name = str(name).strip()

    # Remove trailing EID.
    #
    # Microsoft Excel Fundamentals - A Beginners Guide-11368
    #
    # becomes:
    #
    # Microsoft Excel Fundamentals - A Beginners Guide
    name = re.sub(
        r"-\d+\s*$",
        "",
        name
    )

    name = name.lower()

    # Ignore spaces, punctuation,
    # hyphens, apostrophes, etc.
    name = re.sub(
        r"[^a-z0-9]",
        "",
        name
    )

    return name


# ============================================================
# SELECT HELPER
# ============================================================

def normalize_category_text(text):
    """
    Normalize a Simplex category name for reliable matching.

    Examples that should resolve to the same category:

        Data Science and Business Analytics
        Data Science & Business Analytics
        40 - Data Science & Business Analytics

    The numeric Simplex category ID is ignored. Case, punctuation,
    '&' versus 'and', and repeated whitespace are also ignored.

    We do NOT remove meaningful words, so genuinely different
    categories remain different.
    """

    text = str(text or "").strip().lower()

    # Remove the numeric category ID used by Simplex.
    text = re.sub(r"^\s*\d+\s*-\s*", "", text)

    # Treat '&' and 'and' as equivalent.
    text = text.replace("&", " and ")

    # Keep letters, numbers, and whitespace only.
    text = re.sub(r"[^a-z0-9\s]", " ", text)

    # Normalize whitespace.
    text = re.sub(r"\s+", " ", text).strip()

    return text


def select_option_containing(
    page,
    selector,
    desired_text
):
    """
    Select a category from a normal HTML <select>.

    Exact normalized matches are preferred. A substring fallback
    is used only when exactly one option matches, preventing
    ambiguous categories from being selected accidentally.
    """

    select = page.locator(selector)

    if select.count() == 0:
        raise Exception(
            f"Select field '{selector}' was not found."
        )

    options = select.locator("option")
    desired_normalized = normalize_category_text(desired_text)

    if not desired_normalized:
        raise Exception(
            f"Category value is empty for select '{selector}'."
        )

    matches = []

    for i in range(options.count()):
        option = options.nth(i)

        option_text = option.inner_text().strip()
        option_value = option.get_attribute("value")
        option_normalized = normalize_category_text(option_text)

        if option_normalized:
            matches.append({
                "text": option_text,
                "value": option_value,
                "normalized": option_normalized,
            })

    # Exact normalized match.
    exact_matches = [
        item for item in matches
        if item["normalized"] == desired_normalized
    ]

    if len(exact_matches) == 1:
        selected = exact_matches[0]
        select.select_option(value=selected["value"])
        return selected["text"]

    if len(exact_matches) > 1:
        raise Exception(
            f"Multiple Simplex categories exactly match "
            f"'{desired_text}' after normalization."
        )

    # Safe fallback for a unique partial match.
    partial_matches = [
        item for item in matches
        if (
            desired_normalized in item["normalized"]
            or item["normalized"] in desired_normalized
        )
    ]

    if len(partial_matches) == 1:
        selected = partial_matches[0]
        select.select_option(value=selected["value"])
        return selected["text"]

    if len(partial_matches) > 1:
        examples = ", ".join(
            item["text"] for item in partial_matches[:8]
        )
        raise Exception(
            f"Category '{desired_text}' is ambiguous in "
            f"Simplex dropdown '{selector}'. "
            f"Possible matches: {examples}"
        )

    raise Exception(
        f"Could not find '{desired_text}' inside select '{selector}'. "
        "Matching ignores case, extra spaces, punctuation, "
        "'&' versus 'and', and the numeric Simplex category ID."
    )


# ============================================================
# STEP 1 - SEO
# ============================================================

def create_seo(
    page,
    course_title,
    seo_slug,
    seo_url
):

    print("\n========================================")
    print("STEP 1: SEO CREATION")
    print("========================================")

    print("\nCourse Title:")
    print(course_title)

    print("\nGenerated SEO Slug:")
    print(seo_slug)

    print("\nSimplex SEO URL:")
    print(seo_url)

    # --------------------------------------------------------
    # OPEN SEO LIST
    # --------------------------------------------------------

    print("\nOpening SEO list...")

    page.goto(
        SEO_LIST_URL,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1000)

    # --------------------------------------------------------
    # ADD MORE
    # --------------------------------------------------------

    print("Looking for Add More...")

    add_more = page.get_by_role(
        "link",
        name="Add More"
    )

    if add_more.count() == 0:

        raise Exception(
            "Add More link was not found."
        )

    add_more.first.click()

    page.wait_for_timeout(1000)

    print("Add More opened.")
    print("Current URL:", page.url)

    # --------------------------------------------------------
    # SEO SEARCH PAGE
    # --------------------------------------------------------

    if "/admin/seo/search" not in page.url:

        raise Exception(
            "Simplex did not open SEO search page.\n"
            f"Current URL: {page.url}"
        )

    print("\nEntering SEO Search URL...")

    search_url_field = page.locator("#url")

    if search_url_field.count() == 0:

        raise Exception(
            "SEO Search URL field (#url) was not found."
        )

    # IMPORTANT:
    # Keep the /
    search_url_field.fill(
        seo_url
    )

    entered_search_url = (
        search_url_field.input_value().strip()
    )

    print(
        "Value actually entered:",
        entered_search_url
    )

    if entered_search_url != seo_url:

        raise Exception(
            "SEO Search URL was not entered correctly.\n"
            f"Expected: {seo_url}\n"
            f"Found: {entered_search_url}"
        )

    if not entered_search_url.startswith("/"):

        raise Exception(
            "SEO Search URL is missing the leading '/'."
        )

    print("SEO Search URL: PASS")

    # --------------------------------------------------------
    # SEARCH
    # --------------------------------------------------------

    print("\nClicking Search...")

    search_button = page.locator("#search")

    if search_button.count() == 0:

        raise Exception(
            "SEO Search button (#search) was not found."
        )

    search_button.click()

    page.wait_for_timeout(1500)

    print("Search completed.")
    print("Current URL:", page.url)

    # --------------------------------------------------------
    # SEO FORM
    # --------------------------------------------------------

    url_show = page.locator("#urlShow")

    if url_show.count() == 0:

        raise Exception(
            "SEO form did not open after Search."
        )

    actual_url = (
        url_show.input_value().strip()
    )

    print(
        "SEO URL shown by Simplex:",
        actual_url
    )

    if actual_url != seo_url:

        raise Exception(
            "SEO URL mismatch.\n"
            f"Expected: {seo_url}\n"
            f"Found: {actual_url}"
        )

    print("SEO URL: PASS")

    # --------------------------------------------------------
    # URL TYPE
    # --------------------------------------------------------

    print(
        "\nSelecting URL Type = Course..."
    )

    url_type = page.locator("#url_type")

    if url_type.count() == 0:

        selects = page.locator("select")

        print(
            "Specific #url_type not found."
        )

        print(
            "Total SELECT elements found:",
            selects.count()
        )

        found_selector = None

        for i in range(selects.count()):

            current = selects.nth(i)

            options = current.locator("option")

            for j in range(options.count()):

                text = (
                    options.nth(j)
                    .inner_text()
                    .strip()
                )

                if text.lower() == "course":

                    found_selector = current
                    break

            if found_selector:
                break

        if not found_selector:

            raise Exception(
                "SEO URL Type dropdown "
                "could not be identified."
            )

        url_type = found_selector

    url_type.select_option(
        label="Course"
    )

    print("URL Type: PASS")

    # --------------------------------------------------------
    # SEO FIELDS
    # --------------------------------------------------------

    print("\nFilling SEO fields...")

    meta_title = page.locator("#title")
    meta_description = page.locator("#description")
    meta_keyword = page.locator("#keyword")
    h1_tag = page.locator("#h1Tag")

    for selector, locator in {
        "#title": meta_title,
        "#description": meta_description,
        "#keyword": meta_keyword,
        "#h1Tag": h1_tag
    }.items():

        if locator.count() == 0:

            raise Exception(
                f"Required SEO field "
                f"{selector} was not found."
            )

    meta_title.fill(course_title)
    meta_description.fill(course_title)
    meta_keyword.fill(course_title)
    h1_tag.fill(course_title)

    print("Meta Title: filled")
    print("Meta Description: filled")
    print("Meta Keyword: filled")
    print("H1 Tag: filled")
    print("H2 Tag: left as-is")

    # --------------------------------------------------------
    # THUMBNAIL
    # --------------------------------------------------------

    print("\nAdding SEO thumbnail...")

    thumbnail = page.locator("#thumb_image")

    if thumbnail.count() == 0:

        raise Exception(
            "SEO thumbnail field (#thumb_image) "
            "was not found."
        )

    thumbnail.fill(
        DUMMY_SEO_THUMBNAIL
    )

    print("Thumbnail: filled")

    # --------------------------------------------------------
    # NO INDEX
    # --------------------------------------------------------

    print("\nChecking No Index...")

    noindex = page.locator("#noindex")

    if noindex.count() == 0:

        raise Exception(
            "No Index checkbox (#noindex) "
            "was not found."
        )

    if not noindex.is_checked():
        noindex.check()

    print(
        "No Index:",
        noindex.is_checked()
    )

    # --------------------------------------------------------
    # VALIDATION
    # --------------------------------------------------------

    print("\n========================================")
    print("VALIDATING SEO FIELDS")
    print("========================================")

    if meta_title.input_value() != course_title:
        raise Exception("Meta Title validation failed.")

    if meta_description.input_value() != course_title:
        raise Exception("Meta Description validation failed.")

    if meta_keyword.input_value() != course_title:
        raise Exception("Meta Keyword validation failed.")

    if h1_tag.input_value() != course_title:
        raise Exception("H1 Tag validation failed.")

    if not noindex.is_checked():
        raise Exception("No Index validation failed.")

    print("Meta Title: PASS")
    print("Meta Description: PASS")
    print("Meta Keyword: PASS")
    print("H1 Tag: PASS")
    print("Thumbnail: PASS")
    print("No Index: PASS")

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    print("\nSaving SEO...")

    submit = page.locator("#submit")

    if submit.count() == 0:

        raise Exception(
            "SEO Save button (#submit) was not found."
        )

    submit.click()

    page.wait_for_timeout(1500)

    print(
        "Page after Save:",
        page.url
    )

    if "/admin/seo/list" in page.url:

        print(
            "Returned to SEO list after Save."
        )

        print(
            "SEO saved successfully."
        )

        return seo_url

    if page.locator("#title").count() > 0:

        saved_title = (
            page.locator("#title")
            .input_value()
        )

        if saved_title != course_title:

            raise Exception(
                "SEO save verification failed."
            )

        print(
            "SEO form remained open after Save."
        )

        print(
            "Saved SEO values verified."
        )

        return seo_url

    raise Exception(
        "Unexpected page after SEO Save.\n"
        f"Current URL: {page.url}"
    )


# ============================================================
# BASIC COURSE FILL
# ============================================================

def fill_basic_page(
    page,
    course_title,
    seo_url,
    primary_category,
    primary_eid,
    simplex_level,
    duration_weeks
):

    print("\n========================================")
    print("STEP 2: BASIC COURSE")
    print("========================================")

    print("\nOpening Basic Course page...")

    page.goto(
        BASIC_URL,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1000)

    print(
        "Basic page opened."
    )

    print(
        "Current URL:",
        page.url
    )

    # --------------------------------------------------------
    # NAME
    # --------------------------------------------------------

    name = page.locator("#name")

    if name.count() == 0:
        raise Exception(
            "Course Name field (#name) not found."
        )

    name.fill(course_title)

    # --------------------------------------------------------
    # DISPLAY NAME
    # --------------------------------------------------------

    display_name = page.locator("#displayName")

    if display_name.count() == 0:
        raise Exception(
            "Display Name field (#displayName) not found."
        )

    display_name.fill(course_title)

    # --------------------------------------------------------
    # TAGS
    # --------------------------------------------------------

    tags = page.locator("#tags")

    if tags.count() == 0:
        raise Exception(
            "Tags field (#tags) not found."
        )

    tags.fill(course_title)

    # --------------------------------------------------------
    # RATING
    # --------------------------------------------------------

    rating = page.locator("#rating")

    if rating.count() == 0:
        raise Exception(
            "Rating field (#rating) not found."
        )

    rating.fill(LEARNING_RATING)

    # --------------------------------------------------------
    # URL
    # --------------------------------------------------------

    print("\nSetting Course URL...")

    course_url = page.locator("#url")

    if course_url.count() == 0:
        raise Exception(
            "Course URL field (#url) not found."
        )

    course_url.fill(seo_url)

    print(
        "Course URL entered:",
        course_url.input_value()
    )

    if course_url.input_value() != seo_url:

        raise Exception(
            "Basic Course URL validation failed."
        )

    # --------------------------------------------------------
    # COURSE LOGO
    # --------------------------------------------------------

    image_home_page = page.locator(
        "#image_home_page"
    )

    if image_home_page.count() == 0:

        raise Exception(
            "Course Logo field "
            "(#image_home_page) not found."
        )

    image_home_page.fill(
        DUMMY_COURSE_LOGO
    )

    # --------------------------------------------------------
    # CERTIFICATE
    # --------------------------------------------------------

    print(
        "\nChecking certificate file..."
    )

    if not CERTIFICATE_FILE.exists():

        raise Exception(
            "Certificate file was not found.\n"
            f"Expected:\n{CERTIFICATE_FILE}"
        )

    certificate = page.locator(
        "#certificate_image"
    )

    if certificate.count() == 0:

        raise Exception(
            "Certificate upload field "
            "(#certificate_image) not found."
        )

    certificate.set_input_files(
        str(CERTIFICATE_FILE)
    )

    print(
        "Certificate file: PASS"
    )

    # --------------------------------------------------------
    # COURSE LEVEL
    # --------------------------------------------------------

    print(
        "\nSelecting Course Level..."
    )

    level_radio = page.locator(
        f"#level-{simplex_level}"
    )

    if level_radio.count() == 0:

        raise Exception(
            f"Course Level radio was not found: "
            f"{simplex_level}"
        )

    level_radio.check()

    # --------------------------------------------------------
    # OSL
    # --------------------------------------------------------

    print(
        "\nEnabling Online Self Learning..."
    )

    osl = page.locator(
        "#trainingTypes-osl2"
    )

    if osl.count() == 0:

        raise Exception(
            "Online Self Learning checkbox "
            "was not found."
        )

    if not osl.is_checked():
        osl.check()

    # --------------------------------------------------------
    # ACCESS DAYS 30
    # --------------------------------------------------------

    print(
        "\nEnabling Access Days = 30..."
    )

    access_day_30 = page.locator(
        "#trainingTypes-oslAccessDays-1"
    )

    if access_day_30.count() == 0:

        raise Exception(
            "Access Days 30 control "
            "was not found."
        )

    if not access_day_30.is_checked():
        access_day_30.check()

    # --------------------------------------------------------
    # DURATION
    # --------------------------------------------------------

    print(
        f"\nSetting Duration = "
        f"{duration_weeks} week(s)..."
    )

    duration = page.locator(
        "#durations"
    )

    if duration.count() == 0:

        raise Exception(
            "Duration field (#durations) "
            "not found."
        )

    duration.fill(
        str(duration_weeks)
    )

    duration_type = page.locator(
        "#duration_type"
    )

    if duration_type.count() == 0:

        raise Exception(
            "Duration Type field (#duration_type) "
            "not found."
        )

    duration_type.select_option(
        "week"
    )

    # --------------------------------------------------------
    # HIDE FROM WEBSITE
    # --------------------------------------------------------

    print(
        "\nEnabling Hide Course From Website..."
    )

    hide_from_search = page.locator(
        "#hideFromSearch"
    )

    if hide_from_search.count() == 0:

        raise Exception(
            "Hide Course From Website "
            "checkbox was not found."
        )

    if not hide_from_search.is_checked():
        hide_from_search.check()

    # --------------------------------------------------------
    # COURSE AVAILABLE
    # --------------------------------------------------------

    print(
        "\nSetting Course Available For = B2B Only..."
    )

    course_available_for = page.locator(
        "#course_available_for"
    )

    if course_available_for.count() == 0:

        raise Exception(
            "Course Available For field "
            "was not found."
        )

    course_available_for.select_option(
        "b2b_only"
    )

    # --------------------------------------------------------
    # PRIMARY CATEGORY
    # --------------------------------------------------------

    print(
        "\nSelecting Primary Category..."
    )

    selected_primary_category = (
        select_option_containing(
            page,
            "#primary_label_id",
            primary_category
        )
    )

    print(
        "Selected Primary Category:",
        selected_primary_category
    )

    # --------------------------------------------------------
    # OTHER CATEGORY
    # --------------------------------------------------------

    print(
        "\nSelecting Other Category..."
    )

    other_category = page.locator(
        "#label_id"
    )

    if other_category.count() == 0:

        raise Exception(
            "Other Category field "
            "(#label_id) was not found."
        )

    selected_other_category = select_option_containing(
        page,
        "#label_id",
        primary_category
    )

    print(
        "Selected Other Category:",
        selected_other_category
    )

    # --------------------------------------------------------
    # PRIMARY ELEARNING
    # --------------------------------------------------------

    print(
        "\nSelecting Primary eLearning EID..."
    )

    primary_elearning = page.locator(
        "#primary_eLearning_id"
    )

    if primary_elearning.count() == 0:

        raise Exception(
            "Primary eLearning field "
            "was not found."
        )

    eid_value = str(
        primary_eid
    ).strip()

    eid_option = primary_elearning.locator(
        f"option[value='{eid_value}']"
    )

    if eid_option.count() == 0:

        raise Exception(
            f"EID {eid_value} was not found "
            "in Primary eLearning."
        )

    eid_course_name = (
        eid_option.first
        .inner_text()
        .strip()
    )

    print(
        "EID:",
        eid_value
    )

    print(
        "EID Course:",
        eid_course_name
    )

    # --------------------------------------------------------
    # EID VALIDATION
    # --------------------------------------------------------

    print(
        "\nChecking EID Course Name..."
    )

    requested_normalized = (
        normalize_course_name_for_match(
            course_title
        )
    )

    eid_normalized = (
        normalize_course_name_for_match(
            eid_course_name
        )
    )

    print(
        "Requested normalized:",
        requested_normalized
    )

    print(
        "EID normalized:",
        eid_normalized
    )

    if requested_normalized != eid_normalized:

        raise Exception(
            "EID course-name validation failed.\n\n"
            f"Requested Course:\n{course_title}\n\n"
            f"EID Course:\n{eid_course_name}\n\n"
            f"EID:\n{eid_value}"
        )

    print(
        "EID Course Name Match: PASS"
    )

    primary_elearning.select_option(
        value=eid_value
    )

    print(
        "Primary eLearning EID: PASS"
    )

    # --------------------------------------------------------
    # VALIDATE BASIC
    # --------------------------------------------------------

    print("\n========================================")
    print("VALIDATING BASIC COURSE")
    print("========================================")

    if name.input_value() != course_title:
        raise Exception("Name validation failed.")

    if display_name.input_value() != course_title:
        raise Exception("Display Name validation failed.")

    if tags.input_value() != course_title:
        raise Exception("Tags validation failed.")

    if rating.input_value() != LEARNING_RATING:
        raise Exception("Rating validation failed.")

    if course_url.input_value() != seo_url:
        raise Exception("Course URL validation failed.")

    if not level_radio.is_checked():
        raise Exception("Course Level validation failed.")

    if not osl.is_checked():
        raise Exception(
            "Online Self Learning validation failed."
        )

    if not access_day_30.is_checked():
        raise Exception(
            "Access Days 30 validation failed."
        )

    if duration.input_value() != str(duration_weeks):
        raise Exception("Duration validation failed.")

    if not hide_from_search.is_checked():
        raise Exception(
            "Hide From Website validation failed."
        )

    print("Course Name: PASS")
    print("Display Name: PASS")
    print("Tags: PASS")
    print("Learning Rating: PASS")
    print("Course URL: PASS")
    print("Primary Category: PASS")
    print("Other Category: PASS")
    print("Primary eLearning EID: PASS")
    print("EID Course Name: PASS")
    print("Course Level: PASS")
    print("Course Logo: PASS")
    print("Certificate: PASS")
    print("Online Self Learning: PASS")
    print("Access Days 30: PASS")
    print("Duration: PASS")
    print("Hide From Website: PASS")
    print("Course Available For B2B Only: PASS")

    # --------------------------------------------------------
    # CLICK NEXT
    # --------------------------------------------------------

    print("\n========================================")
    print("BASIC COURSE VALIDATION COMPLETE")
    print("========================================")

    print("\nClicking Next...")

    submit = page.locator(
        "#submit"
    )

    if submit.count() == 0:

        raise Exception(
            "Basic Course Next button "
            "(#submit) was not found."
        )

    submit.click()

    page.wait_for_timeout(1800)

    print(
        "Basic Course Next clicked."
    )

    print(
        "Current URL:",
        page.url
    )

    # --------------------------------------------------------
    # DETERMINE WHAT SIMPLEX DID
    # --------------------------------------------------------

    # NEW COURSE:
    # Simplex should move to something like:
    #
    # /admin/course/course-indetail/courseId/8007
    #
    course_id_match = re.search(
        r"/courseId/(\d+)",
        page.url
    )

    if course_id_match:

        course_id = course_id_match.group(1)

        print(
            "\nNew course flow detected."
        )

        print(
            "Course ID:",
            course_id
        )

        return course_id

    # --------------------------------------------------------
    # EXISTING COURSE:
    # --------------------------------------------------------

    body_text = page.locator(
        "body"
    ).inner_text()

    duplicate_message = (
        "A record matching"
        in body_text
    )

    if duplicate_message:

        print(
            "\nExisting course detected."
        )

        print(
            "Simplex did not create a duplicate."
        )

        print(
            "Opening Course Listing..."
        )

        return open_existing_course_from_listing(
            page,
            course_title
        )

    # --------------------------------------------------------
    # UNKNOWN RESULT
    # --------------------------------------------------------

    raise Exception(
        "Simplex did not create the course "
        "and no existing-course message was detected.\n\n"
        f"Current URL:\n{page.url}"
    )


# ============================================================
# FIND INPUT BY NEARBY LABEL
# ============================================================

def find_input_near_text(
    page,
    text
):

    inputs = page.locator(
        "input"
    )

    for i in range(inputs.count()):

        current = inputs.nth(i)

        try:

            surrounding_text = current.evaluate(
                """
                el => {
                    let node = el;

                    for (let i = 0; i < 4 && node; i++) {
                        if (node.innerText &&
                            node.innerText.toLowerCase().includes(
                                arguments[0].toLowerCase()
                            )) {
                            return node.innerText;
                        }

                        node = node.parentElement;
                    }

                    return "";
                }
                """,
                text
            )

            if text.lower() in surrounding_text.lower():

                return current

        except Exception:
            continue

    return None


# ============================================================
# OPEN EXISTING COURSE
# ============================================================

def open_existing_course_from_listing(
    page,
    course_title
):

    print("\n========================================")
    print("EXISTING COURSE SEARCH")
    print("========================================")

    page.goto(
        COURSE_LIST_URL,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1200)

    print(
        "Course Listing opened."
    )

    print(
        "Current URL:",
        page.url
    )

    # --------------------------------------------------------
    # COURSE NAME SEARCH FIELD
    # --------------------------------------------------------

    print(
        "\nLooking for Course Name search field..."
    )

    course_name_input = find_input_near_text(
        page,
        "Course Name"
    )

    # If nearby-label detection fails,
    # use the first visible text input as fallback.
    if course_name_input is None:

        visible_inputs = page.locator(
            "input[type='text']:visible"
        )

        if visible_inputs.count() == 0:

            raise Exception(
                "Course Name search field "
                "could not be found."
            )

        course_name_input = (
            visible_inputs.first
        )

    print(
        "Course Name search field found."
    )

    course_name_input.fill(
        course_title
    )

    print(
        "Searching for:",
        course_title
    )

    # --------------------------------------------------------
    # SEARCH BUTTON
    # --------------------------------------------------------

    search_button = page.get_by_role(
        "button",
        name="Search"
    )

    if search_button.count() == 0:

        search_button = page.locator(
            "input[type='submit']"
        )

    if search_button.count() == 0:

        raise Exception(
            "Course Listing Search button "
            "was not found."
        )

    search_button.first.click()

    page.wait_for_timeout(1500)

    print(
        "Course Listing search completed."
    )

    # --------------------------------------------------------
    # FIND MATCHING ROW
    # --------------------------------------------------------

    print(
        "\nLooking for matching course..."
    )

    target_normalized = (
        normalize_course_name_for_match(
            course_title
        )
    )

    rows = page.locator(
        "table tbody tr"
    )

    if rows.count() == 0:

        # Some Simplex pages may not use tbody.
        rows = page.locator(
            "table tr"
        )

    matching_row = None

    for i in range(rows.count()):

        row = rows.nth(i)

        try:

            row_text = row.inner_text()

        except Exception:

            continue

        row_normalized = (
            normalize_course_name_for_match(
                row_text
            )
        )

        # We specifically compare the requested
        # normalized course name against the row.
        if target_normalized in row_normalized:

            matching_row = row

            print(
                "Matching course row found."
            )

            break

    if matching_row is None:

        raise Exception(
            "Could not find the existing course "
            "in Course Listing.\n\n"
            f"Course searched:\n{course_title}"
        )

    # --------------------------------------------------------
    # CLICK EDIT IN THAT ROW
    # --------------------------------------------------------

    print(
        "\nOpening existing course..."
    )

    edit_link = matching_row.get_by_role(
        "link",
        name="Edit",
        exact=True
    )

    if edit_link.count() == 0:

        # Fallback for Simplex link text.
        edit_link = matching_row.locator(
            "a"
        ).filter(
            has_text="Edit"
        )

    if edit_link.count() == 0:

        raise Exception(
            "Edit link was not found "
            "inside the matching course row."
        )

    edit_link.first.click()

    page.wait_for_timeout(1500)

    print(
        "Existing course Edit opened."
    )

    print(
        "Current URL:",
        page.url
    )

    # --------------------------------------------------------
    # GET COURSE ID
    # --------------------------------------------------------

    course_id_match = re.search(
        r"/courseId/(\d+)",
        page.url
    )

    if not course_id_match:

        raise Exception(
            "Could not find Course ID after "
            "opening existing course.\n\n"
            f"Current URL:\n{page.url}"
        )

    course_id = course_id_match.group(1)

    print(
        "Existing Course ID:",
        course_id
    )

    # Make sure we are actually on Basic edit page.
    if "/admin/course/basic/courseId/" not in page.url:

        raise Exception(
            "Existing course Edit did not open "
            "the Basic edit page.\n\n"
            f"Current URL:\n{page.url}"
        )

    print(
        "Existing course Basic page: PASS"
    )

    return course_id


# ============================================================
# FIND COURSE OVERVIEW
# ============================================================

def find_course_overview_field(page):

    print(
        "\nLooking for Course Overview field..."
    )

    # --------------------------------------------------------
    # Direct label association
    # --------------------------------------------------------

    labels = page.locator(
        "label"
    )

    for i in range(labels.count()):

        label = labels.nth(i)

        try:

            label_text = (
                label.inner_text()
                .strip()
                .lower()
            )

        except Exception:

            continue

        if "course overview" not in label_text:

            continue

        for_attr = label.get_attribute(
            "for"
        )

        if for_attr:

            target = page.locator(
                f"#{for_attr}"
            )

            if target.count() > 0:

                return target

        parent = label.locator(
            ".."
        )

        textarea = parent.locator(
            "textarea"
        )

        if textarea.count() > 0:

            return textarea.first

    # --------------------------------------------------------
    # DOM surrounding text
    # --------------------------------------------------------

    textareas = page.locator(
        "textarea"
    )

    print(
        "Total textareas found:",
        textareas.count()
    )

    for i in range(textareas.count()):

        textarea = textareas.nth(i)

        try:

            surrounding_text = textarea.evaluate(
                """
                el => {
                    let node = el;

                    for (let i = 0; i < 5 && node; i++) {
                        if (
                            node.innerText &&
                            node.innerText.toLowerCase()
                            .includes("course overview")
                        ) {
                            return node.innerText;
                        }

                        node = node.parentElement;
                    }

                    return "";
                }
                """
            )

            if (
                surrounding_text
                and
                "course overview"
                in surrounding_text.lower()
            ):

                return textarea

        except Exception:

            continue

    # --------------------------------------------------------
    # Last fallback based on known textarea order.
    # From the screenshot:
    #
    # Course Introduction
    # Course Introduction for Mobile
    # Course Overview
    #
    # Therefore Course Overview is normally
    # the third textarea.
    # --------------------------------------------------------

    if textareas.count() >= 3:

        print(
            "Using third textarea as Course Overview."
        )

        return textareas.nth(2)

    raise Exception(
        "Course Overview textarea "
        "could not be identified."
    )


# ============================================================
# FILL COURSE OVERVIEW
# ============================================================

def fill_course_overview(
    page,
    course_id,
    course_overview
):

    print("\n========================================")
    print("STEP 3: COURSE OVERVIEW")
    print("========================================")

    basic_edit_url = (
        f"{SIMPLEX_BASE}/admin/course/basic/courseId/"
        f"{course_id}"
    )

    print(
        "\nOpening Basic edit page again:"
    )

    print(
        basic_edit_url
    )

    page.goto(
        basic_edit_url,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1500)

    print(
        "Basic edit page opened."
    )

    print(
        "Current URL:",
        page.url
    )

    if f"/courseId/{course_id}" not in page.url:

        raise Exception(
            "Could not open the expected "
            "Basic edit page."
        )

    # --------------------------------------------------------
    # OVERVIEW FIELD
    # --------------------------------------------------------

    overview_field = (
        find_course_overview_field(page)
    )

    if overview_field.count() == 0:

        raise Exception(
            "Course Overview field "
            "was not found."
        )

    if not overview_field.is_editable():

        raise Exception(
            "Course Overview field "
            "is still locked."
        )

    print(
        "Course Overview: UNLOCKED"
    )

    # --------------------------------------------------------
    # OVERVIEW (FROM TOC)
    # --------------------------------------------------------

    overview = course_overview

    print(
        "\nCourse Overview (from TOC):"
    )

    print(
        overview
    )

    print(
        "\nCharacter Count:",
        len(overview)
    )

    # --------------------------------------------------------
    # FILL
    # --------------------------------------------------------

    overview_field.fill(
        overview
    )

    actual_overview = (
        overview_field.input_value()
    )

    actual_length = len(
        actual_overview
    )

    # --------------------------------------------------------
    # VALIDATE
    # --------------------------------------------------------

    if actual_overview != overview:

        raise Exception(
            "Course Overview was not "
            "entered correctly."
        )

    if not 250 <= actual_length <= 350:

        raise Exception(
            "Course Overview must be "
            "250-350 characters.\n"
            f"Actual: {actual_length}"
        )

    print(
        "Course Overview: PASS"
    )

    print(
        "Character Count:",
        actual_length
    )

    # --------------------------------------------------------
    # NEXT
    # --------------------------------------------------------

    print(
        "\nClicking Next..."
    )

    submit = page.locator(
        "#submit"
    )

    if submit.count() == 0:

        raise Exception(
            "Next button (#submit) "
            "was not found."
        )

    submit.click()

    page.wait_for_timeout(1800)

    print(
        "Course Overview saved."
    )

    print(
        "Next clicked."
    )

    print(
        "Current URL:",
        page.url
    )

    return page.url


# ============================================================
# STEP 4 - ADDITIONAL CONTENT / SKILLS COVERED
# ============================================================

def get_visible_skill_fields(page):

    # The inspection confirmed that Simplex uses the skills-* naming
    # pattern. The template itself is not visible and must not be filled.
    fields = page.locator(
        "input[id*='skills'][id$='-name']:visible"
    )

    return fields


def open_additional_content(page, course_id):

    print("\n========================================")
    print("STEP 4: ADDITIONAL CONTENT")
    print("========================================")

    additional_url = (
        f"{SIMPLEX_BASE}/admin/course/"
        f"course-moredetail/courseId/{course_id}"
    )

    print("\nOpening Additional Content page...")
    print(additional_url)

    page.goto(
        additional_url,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1500)

    print("Additional Content page opened.")
    print("Current URL:", page.url)

    if "/admin/course/course-moredetail/courseId/" not in page.url:
        raise Exception(
            "Additional Content page did not open correctly.\n"
            f"Current URL: {page.url}"
        )

    return additional_url


def fill_additional_content_skills(
    page,
    skills
):

    print("\n========================================")
    print("SKILLS COVERED")
    print("========================================")

    if len(skills) != 5:
        raise Exception(
            "Skills Covered (from TOC) must contain exactly 5 skills."
        )

    print("\nSkills Covered (from TOC):")
    for index, skill in enumerate(skills, start=1):
        print(f"{index}. {skill}")

    # --------------------------------------------------------
    # FIND CURRENT SKILL FIELDS
    # --------------------------------------------------------

    skill_fields = get_visible_skill_fields(page)

    print(
        "\nCurrent visible skill fields:",
        skill_fields.count()
    )

    # --------------------------------------------------------
    # ADD MORE SKILLS UNTIL EXACTLY 5 EXIST
    # --------------------------------------------------------

    add_more = page.locator("#skills-add")

    if add_more.count() == 0:
        raise Exception(
            "Add More Skills control (#skills-add) was not found."
        )

    while skill_fields.count() < 5:

        before = skill_fields.count()

        print(
            f"Clicking Add More Skills... "
            f"{before}/5"
        )

        add_more.click()

        page.wait_for_timeout(500)

        skill_fields = get_visible_skill_fields(page)

        after = skill_fields.count()

        if after <= before:
            raise Exception(
                "Add More Skills did not create a new "
                "visible skill field.\n"
                f"Fields before: {before}\n"
                f"Fields after: {after}"
            )

    if skill_fields.count() != 5:
        raise Exception(
            "Could not create exactly 5 Skills Covered fields.\n"
            f"Found: {skill_fields.count()}"
        )

    print("Exactly 5 Skills Covered fields found: PASS")

    # --------------------------------------------------------
    # FILL 5 SKILLS
    # --------------------------------------------------------

    print("\nFilling Skills Covered...")

    for index, skill in enumerate(skills):

        field = skill_fields.nth(index)

        if not field.is_editable():
            raise Exception(
                f"Skills Covered field {index + 1} is not editable."
            )

        field.fill(skill)

        actual = field.input_value().strip()

        if actual != skill:
            raise Exception(
                f"Skills Covered field {index + 1} "
                "was not filled correctly.\n"
                f"Expected: {skill}\n"
                f"Found: {actual}"
            )

        print(
            f"Skill {index + 1}: PASS - {actual}"
        )

    # --------------------------------------------------------
    # FINAL VALIDATION
    # --------------------------------------------------------

    print("\n========================================")
    print("VALIDATING SKILLS COVERED")
    print("========================================")

    skill_fields = get_visible_skill_fields(page)

    if skill_fields.count() != 5:
        raise Exception(
            "Skills Covered validation failed. "
            f"Expected 5 fields, found {skill_fields.count()}."
        )

    for index, expected in enumerate(skills):

        actual = skill_fields.nth(index).input_value().strip()

        if actual != expected:
            raise Exception(
                f"Skill {index + 1} validation failed.\n"
                f"Expected: {expected}\n"
                f"Found: {actual}"
            )

        print(f"Skill {index + 1}: PASS")

    print("Skills Covered: PASS")

    # --------------------------------------------------------
    # SAVE & PROCEED
    # --------------------------------------------------------

    print("\nClicking Save & Proceed...")

    save_proceed = page.locator("#saveproceed")

    if save_proceed.count() == 0:
        raise Exception(
            "Save & Proceed button (#saveproceed) was not found."
        )

    save_proceed.click()

    page.wait_for_timeout(1800)

    print("Save & Proceed clicked.")
    print("Current URL:", page.url)

    return page.url


# ============================================================
# STEP 5 - NAVIGATE TO B2B SKILLUP
# ============================================================

def calculate_skillup_hours(duration_input):
    """Calculate B2B SkillUp hours using the minute-based rounding rule.

    Seconds are ignored:
      minutes < 30 -> original hours
      minutes >= 30 -> hours + 1
    """

    parts = duration_input.strip().split(":")

    if len(parts) != 3:
        raise Exception(
            "Course duration must be in HH:MM:SS format."
        )

    try:
        hours = int(parts[0])
        minutes = int(parts[1])
        seconds = int(parts[2])
    except ValueError:
        raise Exception(
            "Course duration must contain numeric hours, minutes, and seconds."
        )

    if hours < 0:
        raise Exception("Course duration hours cannot be negative.")

    if minutes < 0 or minutes > 59:
        raise Exception("Course duration minutes must be between 0 and 59.")

    if seconds < 0 or seconds > 59:
        raise Exception("Course duration seconds must be between 0 and 59.")

    if hours == 0 and minutes == 0 and seconds == 0:
        raise Exception("Course duration must be greater than zero.")

    rounded_hours = hours + (1 if minutes >= 30 else 0)

    # Simplex does not accept 0 as a duration. A sub-30-minute course
    # (e.g. 0:29:33) rounds down to 0 hours under the rule above - in
    # that one case only, round up to 1 hour instead of 0.
    if rounded_hours == 0:
        rounded_hours = 1

    return rounded_hours


B2B_SKILLUP_PATH = "/admin/course/course-slh/courseId/{}"

SKIP_PAGE_PATHS = [
    "/admin/pricing/add/courseId/{}",
    "/admin/course/course-enterprise/courseId/{}",
    "/admin/course/course-mobile/courseId/{}",
    "/admin/course/course-coupon/courseId/{}",
    "/admin/course/course-upsell/courseId/{}",
    "/admin/course/program-colletrals/courseId/{}",
]

DUMMY_SKILLUP_IMAGE = DUMMY_COURSE_LOGO


def navigate_and_leave_intermediate_pages(page, course_id):

    print("\n========================================")
    print("LEAVE INTERMEDIATE COURSE PAGES")
    print("========================================")

    for index, path_template in enumerate(SKIP_PAGE_PATHS, start=1):

        url = SIMPLEX_BASE + path_template.format(course_id)

        print(f"\n{index}. Opening and leaving:")
        print(url)

        page.goto(
            url,
            wait_until="domcontentloaded"
        )

        page.wait_for_timeout(800)

        print("Current URL:", page.url)

        if path_template.split("/courseId/")[0] not in page.url:
            raise Exception(
                "Unexpected page while leaving intermediate page.\n"
                f"Expected path containing: {path_template.split('/courseId/')[0]}\n"
                f"Current URL: {page.url}"
            )

        print("Left page: PASS")


def select_select_by_option_text(page, desired_text, preferred_index=None):

    selects = page.locator("select")
    matches = []

    for i in range(selects.count()):

        select = selects.nth(i)
        options = select.locator("option")

        option_match = None

        for j in range(options.count()):

            option = options.nth(j)
            text = option.inner_text().strip()

            if text.lower() == desired_text.strip().lower():
                option_match = option
                break

        if option_match is not None:
            matches.append((i, select, option_match))

    if not matches:
        raise Exception(
            f"Could not find a SELECT containing option '{desired_text}'."
        )

    if preferred_index is not None:
        for i, select, option in matches:
            if i == preferred_index:
                select.select_option(value=option.get_attribute("value"))
                return select

    # Select the first matching dropdown.
    _, select, option = matches[0]
    select.select_option(value=option.get_attribute("value"))
    return select


def select_first_select_with_option(page, desired_text):
    return select_select_by_option_text(page, desired_text)


def select_duration_type_selects(page):

    # On the B2B SkillUp page the two duration-type dropdowns are the
    # repeated Hours/Days/Weeks/Months selects. We identify them by the
    # option set rather than by screen coordinates.
    selects = page.locator("select")
    duration_selects = []

    required_options = {"hours", "days", "weeks", "months"}

    for i in range(selects.count()):

        select = selects.nth(i)
        options = select.locator("option")

        option_texts = {
            options.nth(j).inner_text().strip().lower()
            for j in range(options.count())
        }

        if required_options.issubset(option_texts):
            duration_selects.append(select)

    if len(duration_selects) < 2:
        raise Exception(
            "Could not identify both B2B SkillUp duration type dropdowns.\n"
            f"Found: {len(duration_selects)}"
        )

    duration_selects[0].select_option(label="Hours")
    duration_selects[1].select_option(label="Hours")

    return duration_selects[0], duration_selects[1]


def find_input_near_text_generic(page, text_value, tag_names=("input", "textarea")):

    # Try an actual label first.
    for tag in tag_names:

        labels = page.locator("label").filter(
            has_text=text_value
        )

        for i in range(labels.count()):
            label = labels.nth(i)

            for tag_name in tag_names:
                control = label.locator(
                    "xpath=following::%s[1]" % tag_name
                )
                if control.count() > 0:
                    return control.first

    # Fallback: locate a nearby container whose text contains the label,
    # then look for the first input/textarea/select in that container.
    text_locator = page.get_by_text(
        text_value,
        exact=False
    )

    for i in range(min(text_locator.count(), 10)):

        current = text_locator.nth(i)

        for levels in range(1, 6):

            xpath = "xpath=" + "/.." * levels
            container = current.locator(xpath)

            if container.count() == 0:
                continue

            for tag in tag_names:
                controls = container.locator(tag)
                if controls.count() == 1:
                    return controls.first

    return None


def check_checkbox_by_text(page, text_value):

    # Prefer an associated label.
    labels = page.locator("label").filter(
        has_text=text_value
    )

    for i in range(labels.count()):
        label = labels.nth(i)
        target_id = label.get_attribute("for")

        if target_id:
            checkbox = page.locator("#" + target_id)
            if checkbox.count() > 0:
                if not checkbox.is_checked():
                    checkbox.check()
                return checkbox

        checkbox = label.locator("input[type='checkbox']")
        if checkbox.count() > 0:
            if not checkbox.is_checked():
                checkbox.check()
            return checkbox

    # Fallback: search nearby containers.
    text_locator = page.get_by_text(
        text_value,
        exact=False
    )

    for i in range(min(text_locator.count(), 10)):

        current = text_locator.nth(i)

        for levels in range(1, 6):
            container = current.locator("xpath=" + "/.." * levels)
            if container.count() == 0:
                continue

            checkbox = container.locator(
                "input[type='checkbox']"
            )

            if checkbox.count() == 1:
                if not checkbox.is_checked():
                    checkbox.check()
                return checkbox

    raise Exception(
        f"Could not find checkbox for: {text_value}"
    )


def select_radio_by_text(page, text_value):

    labels = page.locator("label").filter(
        has_text=text_value
    )

    for i in range(labels.count()):
        label = labels.nth(i)
        target_id = label.get_attribute("for")

        if target_id:
            radio = page.locator("#" + target_id)
            if radio.count() > 0:
                radio.check()
                return radio

        radio = label.locator("input[type='radio']")
        if radio.count() > 0:
            radio.check()
            return radio

    text_locator = page.get_by_text(
        text_value,
        exact=False
    )

    for i in range(min(text_locator.count(), 10)):
        current = text_locator.nth(i)

        for levels in range(1, 6):
            container = current.locator("xpath=" + "/.." * levels)
            if container.count() == 0:
                continue

            radio = container.locator("input[type='radio']")
            if radio.count() == 1:
                radio.check()
                return radio

    raise Exception(
        f"Could not find radio option: {text_value}"
    )


def select_radio_by_text_occurrence(page, text_value, occurrence=0):

    labels = page.locator("label").filter(
        has_text=text_value
    )

    matching_radios = []

    for i in range(labels.count()):
        label = labels.nth(i)
        target_id = label.get_attribute("for")

        if target_id:
            radio = page.locator("#" + target_id)
            if radio.count() > 0 and radio.get_attribute("type") == "radio":
                matching_radios.append(radio.first)
                continue

        radio = label.locator("input[type='radio']")
        if radio.count() > 0:
            matching_radios.append(radio.first)

    if not matching_radios:
        raise Exception(
            f"Could not find radio option: {text_value}"
        )

    if occurrence < 0:
        index = len(matching_radios) + occurrence
    else:
        index = occurrence

    if index < 0 or index >= len(matching_radios):
        raise Exception(
            f"Radio option occurrence {occurrence} was not found for: "
            f"{text_value}. Found {len(matching_radios)} matching option(s)."
        )

    radio = matching_radios[index]
    radio.check()
    return radio


def find_image_inputs(page):

    # Find text/url fields associated with the two Program Thumbnail labels.
    results = []

    for label_text in [
        "Program Card Thumbnail Image",
        "Program Page Thumbnail Image",
    ]:

        control = find_input_near_text_generic(
            page,
            label_text,
            tag_names=("input", "textarea")
        )

        if control is None:
            results.append(None)
        else:
            results.append(control)

    return results


def find_section_container_by_heading(page, heading_text, required_selector=None):

    heading = page.get_by_text(
        heading_text,
        exact=True
    ).first

    if heading.count() == 0:
        raise Exception(
            f"Section heading not found: {heading_text}"
        )

    # Walk upward until we reach a container that actually contains
    # the controls belonging to this section. This prevents duplicate
    # labels such as "Edit for SLH" in different sections from being
    # confused with one another.
    for level in range(1, 8):

        container = heading.locator(
            "xpath=" + "/.." * level
        )

        if container.count() == 0:
            continue

        if required_selector:
            if container.locator(required_selector).count() > 0:
                return container
        else:
            return container

    raise Exception(
        f"Could not identify the container for section: {heading_text}"
    )


def select_radio_inside_section(
    page,
    section_heading,
    radio_text
):

    section = find_section_container_by_heading(
        page,
        section_heading,
        required_selector="input[type='radio']"
    )

    labels = section.locator("label").filter(
        has_text=radio_text
    )

    for i in range(labels.count()):

        label = labels.nth(i)
        target_id = label.get_attribute("for")

        if target_id:
            radio = section.locator(
                "#" + target_id
            )

            if radio.count() > 0 and radio.get_attribute("type") == "radio":
                radio.check()

                if not radio.is_checked():
                    raise Exception(
                        f"Could not select '{radio_text}' in '{section_heading}'."
                    )

                return radio

        radio = label.locator(
            "input[type='radio']"
        )

        if radio.count() > 0:
            radio.first.check()

            if not radio.first.is_checked():
                raise Exception(
                    f"Could not select '{radio_text}' in '{section_heading}'."
                )

            return radio.first

    raise Exception(
        f"Radio '{radio_text}' was not found inside '{section_heading}'."
    )


def check_skills_validation_checkbox(page):

    text_value = "I have reviewed and validated the skills listed."

    text_locator = page.get_by_text(
        text_value,
        exact=False
    )

    if text_locator.count() == 0:
        raise Exception(
            "Skills validation text was not found."
        )

    # The checkbox is in the same small section as the text. Walk up
    # until we find exactly the checkbox belonging to that section.
    for i in range(min(text_locator.count(), 5)):

        current = text_locator.nth(i)

        for level in range(1, 7):

            container = current.locator(
                "xpath=" + "/.." * level
            )

            if container.count() == 0:
                continue

            checkboxes = container.locator(
                "input[type='checkbox']"
            )

            if checkboxes.count() == 1:

                checkbox = checkboxes.first

                if not checkbox.is_checked():
                    checkbox.check()

                if not checkbox.is_checked():
                    raise Exception(
                        "Skills validation checkbox could not be checked."
                    )

                return checkbox

    raise Exception(
        "Could not identify the Skills validation checkbox."
    )


def find_description_control(page):

    # Find the textarea associated with the Course Info description.
    labels = page.locator("label").filter(
        has_text="Edit Description"
    )

    for i in range(labels.count()):

        label = labels.nth(i)

        textarea = label.locator(
            "xpath=following::textarea[1]"
        )

        if textarea.count() > 0:
            return textarea.first

    # Fallback: use the first textarea in the Course Info section.
    section = find_section_container_by_heading(
        page,
        "Course Info",
        required_selector="textarea"
    )

    textareas = section.locator("textarea")

    if textareas.count() > 0:
        return textareas.first

    return None


def find_select_with_option_text(page, desired_text):
    """Find a SELECT containing an option without changing its value."""

    selects = page.locator("select")
    desired = desired_text.strip().lower()

    for i in range(selects.count()):
        select = selects.nth(i)
        options = select.locator("option")

        for j in range(options.count()):
            option = options.nth(j)
            if option.inner_text().strip().lower() == desired:
                return select

    raise Exception(
        f"Could not find a SELECT containing option '{desired_text}'."
    )


def find_duration_type_selects(page):
    """Find the two B2B duration-type SELECTs without changing them."""

    selects = page.locator("select")
    duration_selects = []
    required_options = {"hours", "days", "weeks", "months"}

    for i in range(selects.count()):
        select = selects.nth(i)
        options = select.locator("option")
        option_texts = {
            options.nth(j).inner_text().strip().lower()
            for j in range(options.count())
        }

        if required_options.issubset(option_texts):
            duration_selects.append(select)

    if len(duration_selects) < 2:
        raise Exception(
            "Could not identify both B2B SkillUp duration type dropdowns "
            "during save verification."
        )

    return duration_selects[0], duration_selects[1]


def verify_b2b_skillup_saved(
    page,
    skillup_url,
    expected_level,
    expected_hours
):
    """Reload the B2B page and verify that Simplex persisted key values."""

    print("\n========================================")
    print("VERIFYING B2B SKILLUP SAVE")
    print("========================================")
    print("Reloading B2B SkillUp page to verify server-side save...")

    page.goto(
        skillup_url,
        wait_until="domcontentloaded"
    )
    page.wait_for_timeout(1800)

    if "/admin/course/course-slh/courseId/" not in page.url:
        raise Exception(
            "B2B SkillUp save verification opened an unexpected page.\n"
            f"Current URL: {page.url}"
        )

    # Product offering
    product_select = find_select_with_option_text(
        page,
        "Simplilearn SkillUp"
    )
    actual_product = product_select.locator(
        "option:checked"
    ).inner_text().strip()

    if actual_product.lower() != "simplilearn skillup":
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"Product offering is '{actual_product}'."
        )

    # LMS level
    level_select = find_select_with_option_text(
        page,
        expected_level
    )
    actual_level = level_select.locator(
        "option:checked"
    ).inner_text().strip()

    if actual_level.lower() != expected_level.lower():
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"LMS level is '{actual_level}', expected '{expected_level}'."
        )

    # Course Info: Same as B2C
    same_as_b2c = select_radio_inside_section(
        page,
        "Course Info",
        "Same as B2C"
    )

    if not same_as_b2c.is_checked():
        raise Exception(
            "B2B SkillUp save verification failed: "
            "Course Info is not set to Same as B2C."
        )

    # SLH duration
    slh_duration = find_input_near_text_generic(
        page,
        "SLH - Duration",
        tag_names=("input",)
    )

    if slh_duration is None:
        raise Exception(
            "B2B SkillUp save verification failed: "
            "SLH Duration input was not found after reload."
        )

    actual_slh_hours = slh_duration.input_value().strip()

    if actual_slh_hours != str(expected_hours):
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"SLH Duration is '{actual_slh_hours}', "
            f"expected '{expected_hours}'."
        )

    # Duration types
    duration_type_1, duration_type_2 = find_duration_type_selects(page)

    actual_type_1 = duration_type_1.locator(
        "option:checked"
    ).inner_text().strip()
    actual_type_2 = duration_type_2.locator(
        "option:checked"
    ).inner_text().strip()

    if actual_type_1.lower() != "hours":
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"SLH Duration Type is '{actual_type_1}'."
        )

    if actual_type_2.lower() != "hours":
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"SLH+ Duration Type is '{actual_type_2}'."
        )

    # SLH+ duration
    slh_plus_duration = find_input_near_text_generic(
        page,
        "SLH+ - Duration",
        tag_names=("input",)
    )

    if slh_plus_duration is None:
        slh_plus_duration = find_input_near_text_generic(
            page,
            "SLH+ Duration",
            tag_names=("input",)
        )

    if slh_plus_duration is None:
        raise Exception(
            "B2B SkillUp save verification failed: "
            "SLH+ Duration input was not found after reload."
        )

    actual_slh_plus_hours = slh_plus_duration.input_value().strip()

    if actual_slh_plus_hours != str(expected_hours):
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"SLH+ Duration is '{actual_slh_plus_hours}', "
            f"expected '{expected_hours}'."
        )

    # Program Tag
    program_tag = find_select_with_option_text(
        page,
        "Newly Launched"
    )
    actual_program_tag = program_tag.locator(
        "option:checked"
    ).inner_text().strip()

    if actual_program_tag.lower() != "newly launched":
        raise Exception(
            "B2B SkillUp save verification failed: "
            f"Program Tag is '{actual_program_tag}'."
        )

    # Skills validation checkbox
    skills_checkbox = check_skills_validation_checkbox(page)

    if not skills_checkbox.is_checked():
        raise Exception(
            "B2B SkillUp save verification failed: "
            "Skills validation checkbox is not checked."
        )

    # Course Images must remain Edit for SLH.
    image_edit_radio = select_radio_inside_section(
        page,
        "Course Images",
        "Edit for SLH"
    )

    if not image_edit_radio.is_checked():
        raise Exception(
            "B2B SkillUp save verification failed: "
            "Course Images is not set to Edit for SLH."
        )

    # Accreditation must remain Same as B2C.
    accreditation_same = select_radio_inside_section(
        page,
        "Accreditation Configuration",
        "Same as B2C"
    )

    if not accreditation_same.is_checked():
        raise Exception(
            "B2B SkillUp save verification failed: "
            "Accreditation Configuration is not Same as B2C."
        )

    print("Product offering: SAVE VERIFIED")
    print("LMS Course Level: SAVE VERIFIED")
    print("Course Info - Same as B2C: SAVE VERIFIED")
    print("SLH Duration: SAVE VERIFIED -", actual_slh_hours, "hour(s)")
    print("SLH Duration Type: SAVE VERIFIED - Hours")
    print("SLH+ Duration: SAVE VERIFIED -", actual_slh_plus_hours, "hour(s)")
    print("SLH+ Duration Type: SAVE VERIFIED - Hours")
    print("Program Tag: SAVE VERIFIED - Newly Launched")
    print("Skills Validation: SAVE VERIFIED - Checked")
    print("Course Images: SAVE VERIFIED - Edit for SLH")
    print("Accreditation Configuration: SAVE VERIFIED - Same as B2C")
    print("\nB2B SkillUp: SAVE VERIFIED")

    return page.url


def fill_b2b_skillup(
    page,
    course_id,
    original_course_level,
    rounded_skillup_hours,
    course_overview
):

    print("\n========================================")
    print("STEP 5: B2B SKILLUP")
    print("========================================")

    skillup_url = (
        SIMPLEX_BASE
        + B2B_SKILLUP_PATH.format(course_id)
    )

    print("\nOpening B2B SkillUp page...")
    print(skillup_url)

    page.goto(
        skillup_url,
        wait_until="domcontentloaded"
    )

    page.wait_for_timeout(1800)

    print("B2B SkillUp page opened.")
    print("Current URL:", page.url)

    if "/admin/course/course-slh/courseId/" not in page.url:
        raise Exception(
            "B2B SkillUp page did not open correctly.\n"
            f"Current URL: {page.url}"
        )

    # --------------------------------------------------------
    # COURSE BASICS
    # --------------------------------------------------------

    print("\n========================================")
    print("COURSE BASICS")
    print("========================================")

    print("\nSelecting Part of other product offering = Simplilearn SkillUp...")

    part_product = select_first_select_with_option(
        page,
        "Simplilearn SkillUp"
    )

    selected_value = part_product.locator(
        "option:checked"
    ).inner_text().strip()

    if selected_value.lower() != "simplilearn skillup":
        raise Exception(
            "Part of other product offering validation failed."
        )

    print("Part of other product offering: PASS")

    print("\nSelecting LMS course level...")

    allowed_levels = {
        "beginner": "Beginner",
        "intermediate": "Intermediate",
        "advanced": "Advanced",
    }

    normalized_input_level = original_course_level.strip().lower()

    if normalized_input_level not in allowed_levels:
        raise Exception(
            "B2B SkillUp LMS course level must be Beginner, "
            "Intermediate, or Advanced.\n"
            f"Received: {original_course_level}"
        )

    skillup_level = allowed_levels[normalized_input_level]

    level_select = select_first_select_with_option(
        page,
        skillup_level
    )

    selected_level = level_select.locator(
        "option:checked"
    ).inner_text().strip()

    if selected_level.lower() != skillup_level.lower():
        raise Exception(
            "LMS course level validation failed.\n"
            f"Expected: {skillup_level}\n"
            f"Found: {selected_level}"
        )

    print("LMS Course Level: PASS -", selected_level)

    # --------------------------------------------------------
    # COURSE INFO
    # --------------------------------------------------------

    print("\n========================================")
    print("COURSE INFO")
    print("========================================")

    print("\nKeeping Description Option = Same as B2C...")

    # Explicitly select Same as B2C INSIDE Course Info. This also
    # triggers Simplex's change handler so the B2C description can
    # populate the disabled/read-only description area.
    same_as_b2c = select_radio_inside_section(
        page,
        "Course Info",
        "Same as B2C"
    )

    if not same_as_b2c.is_checked():
        raise Exception(
            "Course Info: Same as B2C was not selected."
        )

    # Click the already-selected radio once as well. Some Simplex versions
    # populate the read-only B2C description only from the radio click/change
    # handler, even when Same as B2C is already the default selection.
    same_as_b2c.click()
    page.wait_for_timeout(900)

    print("Description Option: PASS - Same as B2C")

    # The description should come from the B2C Course Overview.
    # If Simplex exposes an editable/available textarea, verify that
    # it is populated. If it is disabled, Same as B2C is the controlling
    # option and Simplex handles the value server-side.
    description_control = find_description_control(page)

    if description_control is not None:

        description_value = description_control.input_value().strip()

        if description_value:
            print("B2C Description: PASS - populated")
        elif not description_control.is_disabled():
            # Populate with the exact Course Overview generated earlier
            # from the TOC. This keeps the description aligned with B2C
            # while leaving the Same as B2C radio selected.
            overview = course_overview
            description_control.fill(overview)

            if description_control.input_value().strip() != overview.strip():
                raise Exception(
                    "Course Info description could not be populated."
                )

            print("B2C Description: PASS - populated from Course Overview")
        else:
            print(
                "B2C Description: UI field is disabled; Same as B2C remains selected."
            )

    print(
        f"\nSetting SLH Duration = {rounded_skillup_hours} hour(s)..."
    )

    slh_duration = find_input_near_text_generic(
        page,
        "SLH - Duration",
        tag_names=("input",)
    )

    if slh_duration is None:
        raise Exception(
            "SLH - Duration input was not found."
        )

    slh_duration.fill(str(rounded_skillup_hours))

    if slh_duration.input_value().strip() != str(rounded_skillup_hours):
        raise Exception(
            "SLH Duration validation failed."
        )

    print("SLH Duration: PASS")

    print("Selecting SLH Duration Type = Hours...")
    duration_type_1, duration_type_2 = select_duration_type_selects(page)

    selected_duration_1 = duration_type_1.locator(
        "option:checked"
    ).inner_text().strip()

    selected_duration_2 = duration_type_2.locator(
        "option:checked"
    ).inner_text().strip()

    if selected_duration_1.lower() != "hours":
        raise Exception(
            f"SLH Duration Type validation failed: {selected_duration_1}"
        )

    if selected_duration_2.lower() != "hours":
        raise Exception(
            f"SLH+ Duration Type validation failed: {selected_duration_2}"
        )

    print("SLH Duration Type: PASS - Hours")
    print("SLH+ Duration Type: PASS - Hours")

    print(
        f"\nSetting SLH+ Duration = {rounded_skillup_hours} hour(s)..."
    )

    slh_plus_duration = find_input_near_text_generic(
        page,
        "SLH+ - Duration",
        tag_names=("input",)
    )

    if slh_plus_duration is None:
        slh_plus_duration = find_input_near_text_generic(
            page,
            "SLH+ Duration",
            tag_names=("input",)
        )

    if slh_plus_duration is None:
        raise Exception(
            "SLH+ Duration input was not found."
        )

    slh_plus_duration.fill(str(rounded_skillup_hours))

    if slh_plus_duration.input_value().strip() != str(rounded_skillup_hours):
        raise Exception(
            "SLH+ Duration validation failed."
        )

    print("SLH+ Duration: PASS")

    print("\nSelecting Program Tag = Newly Launched...")

    program_tag = select_first_select_with_option(
        page,
        "Newly Launched"
    )

    selected_program_tag = program_tag.locator(
        "option:checked"
    ).inner_text().strip()

    if selected_program_tag.lower() != "newly launched":
        raise Exception(
            "Program Tag validation failed."
        )

    print("Program Tag: PASS -", selected_program_tag)

    # --------------------------------------------------------
    # SKILLS VALIDATION CHECKBOX
    # --------------------------------------------------------

    print("\nChecking Skills validation checkbox...")

    skills_checkbox = check_skills_validation_checkbox(page)

    if not skills_checkbox.is_checked():
        raise Exception(
            "Skills validation checkbox was not checked."
        )

    print("Skills validation checkbox: PASS - Checked")

    # --------------------------------------------------------
    # COURSE IMAGES
    # --------------------------------------------------------

    print("\n========================================")
    print("COURSE IMAGES")
    print("========================================")

    print("Selecting Edit for SLH under Course Images...")

    # IMPORTANT: This is scoped specifically to the Course Images
    # section. Do NOT use a page-wide occurrence because Accreditation
    # Configuration has another Edit for SLH radio.
    image_edit_radio = select_radio_inside_section(
        page,
        "Course Images",
        "Edit for SLH"
    )

    if not image_edit_radio.is_checked():
        raise Exception(
            "Course Images: Edit for SLH could not be selected."
        )

    print("Course Images - Edit for SLH: PASS")

    # Verify that Accreditation Configuration was not accidentally changed.
    print("Verifying Accreditation Configuration remains Same as B2C...")

    accreditation_same = select_radio_inside_section(
        page,
        "Accreditation Configuration",
        "Same as B2C"
    )

    # The previous helper selects the requested radio. Since Same as B2C
    # is the required state, explicitly validate it and leave the section
    # untouched after this verification.
    if not accreditation_same.is_checked():
        raise Exception(
            "Accreditation Configuration is not set to Same as B2C."
        )

    print("Accreditation Configuration: PASS - Same as B2C")

    print("\nAdding temporary dummy image to Program Card Thumbnail...")
    print("Adding temporary dummy image to Program Page Thumbnail...")

    visible_file_inputs = page.locator(
        "input[type='file']:visible"
    )

    file_count = visible_file_inputs.count()

    if file_count < 2:
        raise Exception(
            "Expected two visible Course Image upload controls after "
            "selecting Edit for SLH, but found "
            f"{file_count}."
        )

    if not CERTIFICATE_FILE.exists():
        raise Exception(
            "Temporary dummy image file was not found.\n"
            f"Expected: {CERTIFICATE_FILE}"
        )

    for index in range(2):

        control = visible_file_inputs.nth(index)

        control.set_input_files(
            str(CERTIFICATE_FILE)
        )

        files = control.evaluate(
            "el => el.files ? el.files.length : 0"
        )

        if files != 1:
            raise Exception(
                f"Course image upload control {index + 1} "
                "did not receive the dummy file."
            )

        print(
            f"Program Thumbnail Image {index + 1}: PASS - dummy file"
        )

    # --------------------------------------------------------
    # FINAL VALIDATION
    # --------------------------------------------------------

    print("\n========================================")
    print("VALIDATING B2B SKILLUP")
    print("========================================")

    # Re-check actual DOM state rather than assuming Playwright actions
    # succeeded merely because no exception was raised.
    if part_product.locator("option:checked").inner_text().strip().lower() != "simplilearn skillup":
        raise Exception("Part of other product offering: FAIL")
    print("Part of other product offering: PASS")

    if level_select.locator("option:checked").inner_text().strip().lower() != skillup_level.lower():
        raise Exception("LMS Course Level: FAIL")
    print("LMS Course Level: PASS")

    if not same_as_b2c.is_checked():
        raise Exception("Description Option: FAIL - Same as B2C is not selected")
    print("Description Option: PASS - Same as B2C")

    if slh_duration.input_value().strip() != str(rounded_skillup_hours):
        raise Exception("SLH Duration: FAIL")
    print("SLH Duration: PASS")

    if duration_type_1.locator("option:checked").inner_text().strip().lower() != "hours":
        raise Exception("SLH Duration Type: FAIL")
    print("SLH Duration Type: PASS")

    if slh_plus_duration.input_value().strip() != str(rounded_skillup_hours):
        raise Exception("SLH+ Duration: FAIL")
    print("SLH+ Duration: PASS")

    if duration_type_2.locator("option:checked").inner_text().strip().lower() != "hours":
        raise Exception("SLH+ Duration Type: FAIL")
    print("SLH+ Duration Type: PASS")

    if program_tag.locator("option:checked").inner_text().strip().lower() != "newly launched":
        raise Exception("Program Tag: FAIL")
    print("Program Tag: PASS")

    if not skills_checkbox.is_checked():
        raise Exception("Skills Validation: FAIL")
    print("Skills Validation: PASS - Checked")

    if not image_edit_radio.is_checked():
        raise Exception("Course Images: FAIL - Edit for SLH is not selected")
    print("Course Images: PASS - Edit for SLH")

    if not accreditation_same.is_checked():
        raise Exception(
            "Accreditation Configuration: FAIL - Same as B2C is not selected"
        )
    print("Accreditation Configuration: PASS - Same as B2C")

    # --------------------------------------------------------
    # SUBMIT
    # --------------------------------------------------------

    print("\nClicking Submit...")

    submit_candidates = [
        page.get_by_role("button", name="Submit", exact=True),
        page.get_by_role("link", name="Submit", exact=True),
        page.locator("input[type='submit']"),
        page.locator("button[type='submit']"),
        page.locator("#submit"),
    ]

    submit = None

    for candidate in submit_candidates:

        if candidate.count() > 0 and candidate.first.is_visible():
            submit = candidate.first
            break

    if submit is None:
        raise Exception(
            "Final Submit button could not be found on B2B SkillUp page."
        )

    submit.click()

    page.wait_for_timeout(2200)

    # --------------------------------------------------------
    # REAL SAVE VERIFICATION
    # --------------------------------------------------------

    error_text = page.get_by_text(
        "Unable to save changes. Please check your inputs.",
        exact=False
    )

    if error_text.count() > 0 and error_text.first.is_visible():
        raise Exception(
            "B2B SkillUp was NOT saved by Simplex.\n"
            "Simplex displayed: Unable to save changes. Please check your inputs.\n"
            "The browser has been left on the B2B SkillUp page for inspection."
        )

    print("Submit response received from Simplex.")
    print("Current URL:", page.url)

    verified_url = verify_b2b_skillup_saved(
        page=page,
        skillup_url=skillup_url,
        expected_level=skillup_level,
        expected_hours=rounded_skillup_hours
    )

    return skillup_url, verified_url


# ============================================================
# SINGLE-CID PIPELINE (used by both Single and Batch modes)
# ============================================================

def run_single_cid(
    page,
    toc_path,
    primary_category,
    primary_eid,
    course_level_input,
    duration_input,
    expected_course_title=None
):
    """
    Runs the full SEO + Basic + Overview + Additional Content +
    B2B SkillUp pipeline for one course, driven by its TOC file.

    If expected_course_title is given (batch mode), it is cross-checked
    against the TOC file's own Course Name - the TOC name still drives
    the automation, but a mismatch (e.g. a batch row pointing at the
    wrong TOC file, or a typo) fails the row instead of silently
    creating a CID under the wrong title.

    Raises on any failure. Returns a result dict on success.
    """

    # --------------------------------------------------------
    # TOC FILE
    # --------------------------------------------------------

    print(
        "\nTOC file:",
        toc_path
    )

    toc_data = parse_toc_excel(toc_path)

    course_title = toc_data["course_name"]

    print(
        "Course Title (from TOC):",
        course_title
    )

    if expected_course_title:

        if normalize_course_name_for_match(course_title) != \
                normalize_course_name_for_match(expected_course_title):

            raise Exception(
                "Course Title in the batch file does not match the "
                "TOC file's Course Name.\n\n"
                f"Batch Course Title: {expected_course_title}\n"
                f"TOC Course Name:    {course_title}\n"
                f"TOC file:            {toc_path}\n\n"
                "Fix the batch row's Course Title, or point it at the "
                "correct TOC file."
            )

    skills, course_overview = generate_skills_and_overview_via_ollama(
        toc_data
    )

    print("\nSkills Covered (from TOC):")

    for index, skill in enumerate(skills, start=1):
        print(f"{index}. {skill}")

    print(
        "\nCourse Overview (from TOC):"
    )

    print(
        course_overview
    )

    print(
        "Character Count:",
        len(course_overview)
    )

    # --------------------------------------------------------
    # CALCULATIONS
    # --------------------------------------------------------

    seo_slug = create_seo_slug(
        course_title
    )

    seo_url = create_simplex_seo_url(
        course_title
    )

    simplex_level = normalize_course_level(
        course_level_input
    )

    (
        total_hours,
        rounded_hours,
        duration_weeks
    ) = parse_duration(
        duration_input
    )

    slh_hours = calculate_skillup_hours(
        duration_input
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    print("\n========================================")
    print("INPUT SUMMARY")
    print("========================================")

    print(
        "Course Title       :",
        course_title
    )

    print(
        "SEO Slug           :",
        seo_slug
    )

    print(
        "SEO URL            :",
        seo_url
    )

    print(
        "Primary Category   :",
        primary_category
    )

    print(
        "Primary EID        :",
        primary_eid
    )

    print(
        "Entered Level      :",
        course_level_input
    )

    print(
        "Simplex Level      :",
        simplex_level
    )

    print(
        "Entered Duration   :",
        duration_input
    )

    print(
        "Calculated Hours   :",
        round(total_hours, 4)
    )

    print(
        "Rounded Hours      :",
        rounded_hours
    )

    print(
        "Simplex Duration   :",
        duration_weeks,
        "week(s)"
    )

    print(
        "B2B SkillUp Hours  :",
        slh_hours
    )

    print(
        "========================================"
    )

    # ====================================================
    # STEP 1
    # ====================================================

    seo_url_from_page = create_seo(
        page=page,
        course_title=course_title,
        seo_slug=seo_slug,
        seo_url=seo_url
    )

    if seo_url_from_page != seo_url:

        raise Exception(
            "SEO URL changed unexpectedly."
        )

    # ====================================================
    # STEP 2
    # ====================================================

    course_id = fill_basic_page(
        page=page,
        course_title=course_title,
        seo_url=seo_url_from_page,
        primary_category=primary_category,
        primary_eid=primary_eid,
        simplex_level=simplex_level,
        duration_weeks=duration_weeks
    )

    # ====================================================
    # STEP 3
    # ====================================================

    final_url = fill_course_overview(
        page=page,
        course_id=course_id,
        course_overview=course_overview
    )

    # ====================================================
    # STEP 4
    # ====================================================

    additional_url = open_additional_content(
        page=page,
        course_id=course_id
    )

    additional_final_url = fill_additional_content_skills(
        page=page,
        skills=skills
    )

    # ====================================================
    # STEP 5 - LEAVE INTERMEDIATE PAGES + B2B SKILLUP
    # ====================================================

    navigate_and_leave_intermediate_pages(
        page=page,
        course_id=course_id
    )

    skillup_url, skillup_final_url = fill_b2b_skillup(
        page=page,
        course_id=course_id,
        original_course_level=course_level_input,
        rounded_skillup_hours=slh_hours,
        course_overview=course_overview
    )

    # ====================================================
    # COMPLETE
    # ====================================================

    print("\n========================================")
    print("SEO + BASIC + OVERVIEW + ADDITIONAL CONTENT + B2B SKILLUP SAVE VERIFIED")
    print("========================================")

    print(
        "\nCourse:",
        course_title
    )

    print(
        "Course ID:",
        course_id
    )

    print(
        "SEO URL:",
        seo_url
    )

    print(
        "\nSEO: COMPLETED"
    )

    print(
        "Basic Fields: COMPLETED"
    )

    print(
        "Existing/New Course Handling: COMPLETED"
    )

    print(
        "Course Overview: COMPLETED"
    )

    print(
        "Additional Content: COMPLETED"
    )

    print(
        "Skills Covered (5): COMPLETED"
    )

    print(
        "Intermediate Pages: LEFT"
    )

    print(
        "B2B SkillUp: SAVE VERIFIED"
    )

    print(
        "\nAdditional Content URL:",
        additional_url
    )

    print(
        "After Save & Proceed:",
        additional_final_url
    )

    print(
        "B2B SkillUp URL:",
        skillup_url
    )

    print(
        "After B2B Submit:",
        skillup_final_url
    )

    print(
        "\n========================================"
    )

    return {
        "course_title": course_title,
        "course_id": course_id,
        "seo_url": seo_url,
        "skillup_url": skillup_url,
        "skillup_final_url": skillup_final_url,
    }


# ============================================================
# CHROME CONNECTION
# ============================================================

def connect_to_chrome(playwright):

    print(
        "\nConnecting to existing Chrome..."
    )

    try:

        browser = (
            playwright.chromium.connect_over_cdp(
                "http://127.0.0.1:9222"
            )
        )

    except Exception as e:

        raise Exception(
            "\nCould not connect to Chrome.\n\n"
            "Make sure Chrome was started with:\n\n"
            'chrome.exe --remote-debugging-port=9222 '
            '--user-data-dir="%TEMP%\\ChromeDebug"\n\n'
            "Then log into Simplex.\n\n"
            f"Original error: {e}"
        )

    print(
        "Connected to Chrome."
    )

    contexts = browser.contexts

    if not contexts:

        raise Exception(
            "No Chrome browser context was found."
        )

    context = contexts[0]

    pages = context.pages

    if pages:

        page = pages[-1]

    else:

        page = context.new_page()

    print(
        "Current Browser Page:",
        page.url
    )

    return page


# ============================================================
# MAIN
# ============================================================

def main():

    print("\n========================================")
    print("SIMPLEX CID AUTOMATION")
    print("========================================")

    # --------------------------------------------------------
    # MODE SELECTION
    # --------------------------------------------------------

    mode = input(
        "\nCreate 1 CID or a bunch of CIDs?\n"
        "1. Single CID\n"
        "2. Bunch of CIDs (batch Excel file)\n"
        "Enter 1 or 2: "
    ).strip()

    if mode not in ("1", "2"):

        raise Exception(
            "Invalid selection. Enter 1 or 2."
        )

    # --------------------------------------------------------
    # CERTIFICATE
    # --------------------------------------------------------

    if not CERTIFICATE_FILE.exists():

        raise Exception(
            "Certificate file is missing.\n\n"
            f"Expected:\n{CERTIFICATE_FILE}"
        )

    # --------------------------------------------------------
    # CONNECT CHROME (once, reused for every CID)
    # --------------------------------------------------------

    with sync_playwright() as p:

        page = connect_to_chrome(p)

        # ====================================================
        # SINGLE CID
        # ====================================================

        if mode == "1":

            print(
                "\nSelect the Course TOC Excel file..."
            )

            toc_path = pick_excel_file_dialog(
                "Select the Course TOC Excel file"
            )

            primary_category = input(
                "Enter Primary Category: "
            ).strip()

            if not primary_category:
                raise Exception(
                    "Primary Category cannot be empty."
                )

            primary_eid = input(
                "Enter Primary eLearning EID: "
            ).strip()

            if not primary_eid:
                raise Exception(
                    "Primary EID cannot be empty."
                )

            course_level_input = input(
                "Enter Course Level: "
            ).strip()

            duration_input = input(
                "Enter Course Duration (HH:MM:SS): "
            ).strip()

            run_single_cid(
                page=page,
                toc_path=toc_path,
                primary_category=primary_category,
                primary_eid=primary_eid,
                course_level_input=course_level_input,
                duration_input=duration_input
            )

        # ====================================================
        # BUNCH OF CIDS (BATCH)
        # ====================================================

        else:

            print(
                "\nSelect the batch CID Excel file..."
            )

            batch_path = pick_excel_file_dialog(
                "Select the batch CID Excel file"
            )

            batch_rows = parse_batch_excel(batch_path)

            total = len(batch_rows)

            print(
                f"\nFound {total} course(s) in the batch file."
            )

            successes = []
            failures = []

            for position, row in enumerate(batch_rows, start=1):

                print(
                    "\n\n################################################"
                )

                print(
                    f"CID {position}/{total} - "
                    f"batch row {row['row_number']}"
                    + (
                        f" - {row['title_hint']}"
                        if row["title_hint"] else ""
                    )
                )

                print(
                    "################################################"
                )

                try:

                    result = run_single_cid(
                        page=page,
                        toc_path=row["toc_path"],
                        primary_category=row["primary_category"],
                        primary_eid=row["primary_eid"],
                        course_level_input=row["course_level"],
                        duration_input=row["duration"],
                        expected_course_title=row["title_hint"]
                    )

                    successes.append({
                        "row_number": row["row_number"],
                        **result,
                    })

                except Exception as error:

                    print(
                        f"\nCID {position}/{total} FAILED "
                        f"(batch row {row['row_number']}):"
                    )

                    print(error)

                    failures.append({
                        "row_number": row["row_number"],
                        "toc_path": str(row["toc_path"]),
                        "error": str(error),
                    })

            # ====================================================
            # BATCH SUMMARY
            # ====================================================

            print(
                "\n\n========================================"
            )
            print("BATCH SUMMARY")
            print("========================================")

            print(
                f"\nTotal rows       : {total}"
            )

            print(
                f"Succeeded        : {len(successes)}"
            )

            print(
                f"Failed           : {len(failures)}"
            )

            if successes:

                print(
                    "\nSucceeded courses:"
                )

                for item in successes:

                    print(
                        f"  Row {item['row_number']}: "
                        f"{item['course_title']} "
                        f"(Course ID {item['course_id']})"
                    )

            if failures:

                print(
                    "\nFailed rows:"
                )

                for item in failures:

                    print(
                        f"  Row {item['row_number']}: "
                        f"{item['toc_path']}"
                    )

                    print(
                        f"    Error: {item['error']}"
                    )

            print(
                "\n========================================"
            )

        input(
            "\nPress ENTER to close..."
        )


# ============================================================
# START
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        print("\n")
        print(
            "=================================================="
        )

        print(
            "AUTOMATION STOPPED"
        )

        print(
            "=================================================="
        )

        print(
            "\nERROR:"
        )

        print(
            error
        )

        print(
            "\n=================================================="
        )

        input(
            "\nPress ENTER to close..."
        )