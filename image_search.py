"""Unsplash图片搜索模块 - 根据关键词推荐配图"""

import requests
from typing import List, Optional, Dict
from io import BytesIO


class UnsplashSearcher:
    """Unsplash图片搜索器"""

    BASE_URL = "https://api.unsplash.com"

    def __init__(self, access_key: str = ""):
        self.access_key = access_key

    def search_images(
        self,
        query: str,
        per_page: int = 3,
        orientation: str = "landscape",
    ) -> List[Dict]:
        """搜索图片

        Args:
            query: 搜索关键词
            per_page: 返回数量(1-30)
            orientation: 方向 landscape/portrait/squarish

        Returns:
            图片信息列表: [{url, width, height, description, download_url}, ...]
        """
        if not self.access_key:
            return []

        try:
            resp = requests.get(
                f"{self.BASE_URL}/search/photos",
                params={
                    "query": query,
                    "per_page": per_page,
                    "orientation": orientation,
                    "client_id": self.access_key,
                },
                timeout=10,
            )
            resp.raise_for_status()
            data = resp.json()

            results = []
            for item in data.get("results", []):
                results.append({
                    "url": item["urls"]["regular"],
                    "thumb_url": item["urls"]["small"],
                    "width": item["width"],
                    "height": item["height"],
                    "description": item.get("description") or item.get("alt_description") or "",
                    "download_url": item["links"]["download"],
                })
            return results

        except Exception as e:
            print(f"Unsplash搜索失败: {e}")
            return []

    def download_image(self, url: str) -> Optional[bytes]:
        """下载图片，返回二进制数据"""
        try:
            resp = requests.get(url, timeout=15)
            resp.raise_for_status()
            return resp.content
        except Exception as e:
            print(f"图片下载失败: {e}")
            return None

    def search_and_download(
        self,
        query: str,
        per_page: int = 1,
        orientation: str = "landscape",
    ) -> List[bytes]:
        """搜索并下载图片

        Returns:
            图片二进制数据列表
        """
        results = self.search_images(query, per_page, orientation)
        images = []
        for item in results:
            data = self.download_image(item["url"])
            if data:
                images.append(data)
        return images


# 关键词到Unsplash搜索词的映射（中→英）
_KEYWORD_MAP = {
    "科技": "technology",
    "教育": "education",
    "自然": "nature",
    "商业": "business",
    "医疗": "medical",
    "艺术": "art",
    "音乐": "music",
    "建筑": "architecture",
    "旅行": "travel",
    "美食": "food",
    "运动": "sports",
    "历史": "history",
    "文化": "culture",
    "数据": "data",
    "团队": "teamwork",
    "创新": "innovation",
    "未来": "future",
    "城市": "city",
    "海洋": "ocean",
    "山": "mountain",
    "花": "flower",
    "天空": "sky",
    "森林": "forest",
    "宇宙": "space",
    "人工智能": "artificial intelligence",
    "机器学习": "machine learning",
    "互联网": "internet",
    "设计": "design",
    "环保": "environment",
    "健康": "health",
}


def translate_keywords(keywords: List[str]) -> str:
    """将中文关键词翻译为英文搜索词

    优先使用映射表，否则直接使用原词（Unsplash也支持部分中文搜索）
    """
    translated = []
    for kw in keywords:
        translated.append(_KEYWORD_MAP.get(kw, kw))
    return " ".join(translated) if translated else ""
