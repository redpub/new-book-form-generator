"""
Field configuration for the book-form generator.

Each section maps template.docx placeholders to:
  - label       Traditional Chinese label shown in the form
  - source_key  Dot-notation key in the AI-extracted flat dict (leave "" to skip auto-fill)
  - type        "text" | "textarea" | "checkbox"
  - hint        Optional placeholder / helper text

Edit this file to correct labels or source mappings without touching any code.

Special source_key values understood by the app:
  ""                  Leave blank; user fills manually
  "extras.<key>"      Pulls from the extras sub-dict returned by the AI
  "#trim_width"       Parses the width part from trim_size (e.g. "150mm X 210mm" → "150")
  "#trim_height"      Parses the height part from trim_size (e.g. "150mm X 210mm" → "210")
"""

# Value written to the DOCX template when a checkbox is ticked / unticked.
CHECKBOX_CHAR = "■"
UNCHECKED_CHAR = "□"

SECTIONS: list[dict] = [
    # ── Page 1 ──────────────────────────────────────────────────────────────
    {
        "title": "書籍基本資料",
        "layout": "grid",
        "fields": [
            {"placeholder": "bookTitle",   "label": "書名",
                "source_key": "title",            "type": "text"},
            {"placeholder": "seriesName",  "label": "系列名",
                "source_key": "series",           "type": "text"},
            {"placeholder": "publisher",   "label": "出版社",
                "source_key": "extras.publisher", "type": "text"},
            {"placeholder": "author",      "label": "作者",
                "source_key": "author_name",      "type": "text"},
            {"placeholder": "translator",  "label": "譯者（如適用）",
                "source_key": "contributor",      "type": "text"},
            {"placeholder": "isbn",        "label": "ISBN",
                "source_key": "isbn",             "type": "text"},
            {"placeholder": "priceHKD",    "label": "定價（港幣）",
                "source_key": "price",            "type": "text"},
            {"placeholder": "bookCover",   "label": "封面設計",
                "source_key": "",                 "type": "text"},
        ],
    },
    {
        "title": "出版規格",
        "layout": "grid",
        "fields": [
            {"placeholder": "publishDate", "label": "出版日期 (dd/mm/yyyy)",
                "source_key": "publication_date", "type": "date"},
            {"placeholder": "moFormat",    "label": "開度",
                "source_key": "",                 "type": "text",
                "options": ["", "16", "25", "32", "64"]},
            {"placeholder": "bookBinding", "label": "裝幀",
                "source_key": "extras.binding",   "type": "text"},
            {"placeholder": "width",       "label": "長 (mm)",
                "source_key": "#trim_width",      "type": "text", "hint": "例：150 (mm，輸出轉為 cm)"},
            {"placeholder": "height",      "label": "闊 (mm)",
                "source_key": "#trim_height",     "type": "text", "hint": "例：210 (mm，輸出轉為 cm)"},
            {"placeholder": "thickness",   "label": "高 (mm)",
                "source_key": "",                 "type": "text", "hint": "輸入 mm，輸出轉為 cm"},
            {"placeholder": "weight",      "label": "重量 (g)",
                "source_key": "",                 "type": "text"},
            {"placeholder": "numOfPages",  "label": "頁數",
                "source_key": "page_count",       "type": "text"},
        ],
    },
    {
        "title": "版次及印刷用色",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "firstEdition",   "label": "初版",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "rePrint",        "label": "再版",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "revisedEdition", "label": "改版",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "bookColor",      "label": "印刷用色",
                "source_key": "extras.print_color", "type": "text"},
        ],
    },
    {
        "title": "附加品",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "accessoryCD",          "label": "CD ROM",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryAudioCD",     "label": "Audio CD",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryMP3CD",       "label": "MP3格式CD",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryDVD",         "label": "DVD",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryGiftCoupon",  "label": "贈品／贈券",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryGiftCouponText",
                "label": "贈品／贈券說明", "source_key": "", "type": "text"},
            {"placeholder": "accessoryOther",       "label": "其他",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "accessoryOtherText",
                "label": "附件其他說明",  "source_key": "", "type": "text"},
        ],
    },
    {
        "title": "語種",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "langChi",     "label": "中文",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "langTradChi", "label": "繁體",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "langSimChi",  "label": "簡體",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "langEng",     "label": "英文",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "langChiEng",  "label": "中英對照",
                "source_key": "", "type": "checkbox"},
        ],
    },
    {
        "title": "版權地區",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "copyrightRegionHongKong", "label": "香港",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "copyrightRegionMacau",    "label": "澳門",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "copyrightRegionChina",    "label": "中國",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "copyrightRegionTaiwan",   "label": "台灣",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "copyrightRegionOther",    "label": "其他地方",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "copyrightRegionOtherText",
                "label": "其他地方說明", "source_key": "", "type": "text"},
        ],
    },
    {
        "title": "目標市場分析",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "",        "label": "1. 年齡",
                "type": "label", "source_key": ""},
            {"placeholder": "fromAge", "label": "從（歲）",
                "source_key": "", "type": "text"},
            {"placeholder": "toAge",   "label": "至（歲）",
                "source_key": "", "type": "text"},
            {"placeholder": "",             "label": "2. 性別",
                "type": "label", "source_key": ""},
            {"placeholder": "genderMale",   "label": "男",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "genderFemale", "label": "女",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "", "label": "3. 讀者對象",
                "type": "label", "source_key": ""},
            {"placeholder": "audienceInfants",
                "label": "幼兒",          "source_key": "", "type": "checkbox"},
            {"placeholder": "audiencePrimarySchoolStudents",
                "label": "小學生",        "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceSecondaryUniversityStudents",
                "label": "中、大學生",    "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceParents",
                "label": "家長",          "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceWorks",
                "label": "上班族",        "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceGeneralPublic",
                "label": "一般大眾",      "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceProfessional",
                "label": "專業人士",      "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceOther",
                "label": "其他及補充",     "source_key": "", "type": "checkbox"},
            {"placeholder": "audienceOtherText",
                "label": "讀者對象其他說明", "source_key": "", "type": "text"},
            {"placeholder": "buyAudience",
                "label": "4. 購買對象（如有別於讀者對象）",
                "source_key": "target_audience", "type": "text"},
        ],
    },
    {
        "title": "建議銷售渠道",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "salesChannelChiBookStore",
                "label": "中文書店",      "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelEngBookStore",
                "label": "英文書店",      "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannel2ndFloorBookStore",
                "label": "二樓書店",      "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelTextBookStore",
                "label": "教科書書店",    "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelDepartmentStore",
                "label": "百貨公司",      "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelSchool",
                "label": "學校直銷",      "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelConvenienceStore",
                "label": "便利店",        "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelProfessionalAssociation",
                "label": "專業團體直銷",  "source_key": "", "type": "checkbox"},
            {"placeholder": "salesChannelLibrary",
                "label": "圖書館",        "source_key": "", "type": "checkbox"},
        ],
    },
    {
        "title": "相關書籍",
        "layout": "full",
        "fields": [
            {"placeholder": "relatedBookTitle", "label": "相關書名",
                "source_key": "", "type": "textarea"},
        ],
    },
    # ── Page 2 ──────────────────────────────────────────────────────────────
    {
        "title": "內容介紹",
        "layout": "full",
        "fields": [
            {"placeholder": "bookIntro", "label": "內容介紹（請以點列形式列出）",
                "source_key": "synopsis", "type": "html"},
        ],
    },
    {
        "title": "本書特色及賣點",
        "layout": "full",
        "fields": [
            {"placeholder": "bookHighlights", "label": "本書特色及賣點（請以點列形式列出）",
                "source_key": "", "type": "html"},
        ],
    },
    {
        "title": "作者介紹",
        "layout": "full",
        "fields": [
            {"placeholder": "authorIntro", "label": "作者介紹",
                "source_key": "author_bio", "type": "html"},
        ],
    },
    {
        "title": "現有市場對手",
        "layout": "grid",
        "fields": [
            {"placeholder": "competitorPublisher",    "label": "出版社",
                "source_key": "", "type": "text"},
            {"placeholder": "competitorBookTitle",    "label": "書名",
                "source_key": "", "type": "text"},
            {"placeholder": "competitorBookPriceHKD", "label": "定價或特價（港幣）",
                "source_key": "", "type": "text"},
        ],
    },
    {
        "title": "營銷／市場推廣計劃",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "marketingLevel1", "label": "一級",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "marketingLevel2", "label": "二級",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "marketingLevel3", "label": "三級",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "marketingCampaign", "label": "預期之推廣計劃",
                "source_key": "", "type": "textarea"},
        ],
    },
    {
        "title": "推廣品提供",
        "layout": "checkboxes",
        "fields": [
            {"placeholder": "promotionPoster",      "label": "海報",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionLeaflet",     "label": "宣傳單張",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionLuckyBox",    "label": "吉盒",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionDisplayItem", "label": "陳列品",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionPromoter",    "label": "推廣員（需另商議）",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionOther",       "label": "其他",
                "source_key": "", "type": "checkbox"},
            {"placeholder": "promotionOtherText",
                "label": "推廣品其他說明", "source_key": "", "type": "text"},
        ],
    },
]
