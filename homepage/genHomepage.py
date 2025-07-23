import json

with open('homepage_config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

def generate_button_html(button, is_modal=False):
    icon = button.get("icon", "")
    label = button.get("label", "")
    link = button.get("link", "#")
    onclick = f'onclick="{button["onclick"]}"' if "onclick" in button else ""

    # 判斷 bg 和 text 是 class 還是 hex code
    bg_color = button.get("bg_color", "")
    text_color = button.get("text_color", "")

    # 若是 # 開頭，則轉為 inline-style
    styles = []
    if bg_color.startswith("#"):
        styles.append(f"background-color: {bg_color}")
    if text_color.startswith("#") or text_color.lower() in ["white", "black", "red", "pink"]:
        styles.append(f"color: {text_color}")
    style_attr = f'style="{"; ".join(styles)}"' if styles else ""

    # Bootstrap class 直接用
    classes = "modal-button" if is_modal else "content-button"
    if not bg_color.startswith("#"):
        classes += f" {bg_color}"
    if not text_color.startswith("#") and not text_color.lower() in ["white", "black", "red", "pink"]:
        classes += f" {text_color}"

    return f'''
        <a href="{link}" class="{classes.strip()}" {onclick} {style_attr}>
            <i class="bi {icon}"></i>
            <span>{label}</span>
        </a>'''

def generate_tab_button_html(tab, index):
    active_class = "active" if index == 0 else ""
    return f'''
        <li class="nav-item" role="presentation">
            <button class="nav-link {active_class}" id="{tab["id"]}-tab" data-bs-toggle="tab" data-bs-target="#{tab["id"]}" type="button" role="tab">
                <i class="bi {tab["icon"]}"></i> {tab["title"]}
            </button>
        </li>'''

def generate_tab_content_html(tab, index):
    active_class = "show active" if index == 0 else ""
    buttons_html = "\n".join(generate_button_html(btn) for btn in tab["buttons"])
    return f'''
        <div class="tab-pane fade {active_class}" id="{tab["id"]}" role="tabpanel">
            <h3><i class="bi {tab["icon"]}"></i> {tab["heading"]}</h3>
            <p>{tab["description"]}</p>
            <div class="content-grid">
                {buttons_html}
            </div>
        </div>'''

def generate_modal_html(modal):
    buttons_html = "\n".join(generate_button_html(btn, is_modal=True) for btn in modal["buttons"])
    return f'''
<div class="modal fade" id="{modal["id"]}" tabindex="-1" aria-labelledby="{modal["id"]}Label" aria-hidden="true">
    <div class="modal-dialog modal-lg">
        <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="{modal["id"]}Label">
                    <i class="bi bi-printer"></i> {modal["title"]}
                </h5>
                <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
            </div>
            <div class="modal-body">
                <p class="text-muted">{modal["description"]}</p>
                <div class="modal-grid">
                    {buttons_html}
                </div>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">取消</button>
            </div>
        </div>
    </div>
</div>
'''

tab_nav = "\n".join(generate_tab_button_html(tab, i) for i, tab in enumerate(config["tabs"]))
tab_contents = "\n".join(generate_tab_content_html(tab, i) for i, tab in enumerate(config["tabs"]))
modals = "\n".join(generate_modal_html(modal) for modal in config.get("modals", []))

html_output = f'''
<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>網頁入口頁面</title>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-icons/1.10.0/font/bootstrap-icons.min.css" rel="stylesheet">
    <link href="./hompage_component/homepage.css" rel="stylesheet">
</head>
<body>
    <div class="header">
        <h1>系統入口頁面</h1>
    </div>

    <div class="container-fluid">
        <div class="tab-container">
            <ul class="nav nav-tabs" id="mainTabs" role="tablist">
                {tab_nav}
            </ul>

            <div class="tab-content" id="mainTabContent">
                {tab_contents}
            </div>
        </div>
    </div>

    {modals}

'''
modal_scripts = ""
for modal in config.get("modals", []):
    modal_id = modal["id"]
    modal_func = f"show{modal_id[0].upper() + modal_id[1:]}"

    modal_scripts += f"""
function {modal_func}(e) {{
    e.preventDefault();
    const modal = new bootstrap.Modal(document.getElementById('{modal_id}'));
    modal.show();
}}
"""

html_output += f"""
<script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap/5.3.0/js/bootstrap.bundle.min.js"></script>
<script>
{modal_scripts.strip()}
</script>
</body>
</html>
"""

with open('index.html', 'w', encoding='utf-8') as f:
    f.write(html_output)

print("✅ HTML 已成功產生為 index.html")
