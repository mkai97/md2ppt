import re
import json
import markdown2


def markdown_to_json(markdown_text):
    # 将 markdown 文本转换为 HTML
    html_text = markdown2.markdown(markdown_text)

    # 提取第一级标题作为 presentation 的 title
    title_match = re.search(r'<h1>(.*?)</h1>', html_text)
    if title_match:
        presentation_title = title_match.group(1)
    else:
        presentation_title = ""

    # 提取第二级标题及其内容
    slides = []
    slide_matches = re.findall(r'<h2>(.*?)</h2>(.*?)(?=<h2>|</body>)', html_text, re.DOTALL)
    for index, slide_match in enumerate(slide_matches):
        slide_title = slide_match[0]
        slide_content = slide_match[1].strip()
        slide_content_list = re.findall(r'<p>(.*?)</p>', slide_content, re.DOTALL)
        slide = {
            "layout": index + 1,
            "title": slide_title,
            "content": slide_content_list,
            "image": {
                "path": "./md2ppt/WechatIMG35998.jpg",
                "size": {
                    "width": 10,
                    "height": 7.5
                }
            }
        }
        slides.append(slide)

    # 构建 JSON 对象
    json_obj = {
        "presentation": {
            "title": presentation_title,
            "image": {
                "path": "./md2ppt/WechatIMG35998.jpg",
                "size": {
                    "width": 10,
                    "height": 7.5
                }
            },
            "slides": slides
        }
    }

    return json.dumps(json_obj, ensure_ascii=False, indent=2)


# 测试代码
markdown_text = """
# 2024年度雷军演讲

尊敬的听众，大家好！我是雷军，非常荣幸能在这里与大家分享我的一些想法和见解。

小米科技创始人，致力于推动科技与创新的融合。

今天，我将与大家探讨科技行业的发展趋势，以及小米科技在这一过程中的角色和贡献。

## 勇气从何而来

尊敬的听众，大家好！我是雷军，非常荣幸能在这里与大家分享我的一些想法和见解。
小米科技创始人，致力于推动科技与创新的融合。
今天，我将与大家探讨科技行业的发展趋势，以及小米科技在这一过程中的角色和贡献。
"""

json_result = markdown_to_json(markdown_text)
print(json_result)
