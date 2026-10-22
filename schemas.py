# -*- coding: utf-8 -*-
"""行程规划的结构化 Schema（DESIGN §7）。
可选字段一律给默认值，方便 _structured() 构造 fallback 对象、也方便离线测试。
容器说明：with_structured_output 只能绑单个顶层模型、不能绑 list[X]，故大纲用 TripOutline 承载。"""
from __future__ import annotations
from typing import List, Optional, Literal
from pydantic import BaseModel, Field


class TripSpec(BaseModel):
    destination: str = Field(description="目的地城市/区域")
    days: int = Field(description="行程天数（>=1）")
    origin: Optional[str] = Field(default=None, description="出发地，可空")
    budget_cny: Optional[int] = Field(default=None, description="总预算(人民币元)，可空，别编造")
    interests: List[str] = Field(default_factory=list, description="兴趣标签，如 历史/美食/自然")
    pace: Literal["轻松", "适中", "紧凑"] = Field(default="适中", description="节奏")
    party: str = Field(default="1人", description="同行人群，如 情侣/亲子/1人")
    start_date: Optional[str] = Field(default=None, description="出发日期 YYYY-MM-DD，可空")


class DayOutline(BaseModel):
    day: int = Field(description="第几天，从 1 开始")
    theme: str = Field(default="", description="当天主题")
    area: str = Field(default="", description="主要活动区域")
    candidate_spots: List[str] = Field(default_factory=list, description="候选景点名（3~5个）")


class TripOutline(BaseModel):
    """薄容器：承载整份逐日大纲（with_structured_output 需要单个顶层模型）。"""
    days: List[DayOutline] = Field(default_factory=list, description="逐日大纲")


class Slot(BaseModel):
    time: str = Field(description="时段，如 上午/12:00/傍晚")
    activity: str = Field(description="做什么")
    spot: str = Field(default="", description="地点/景点")
    transit: str = Field(default="", description="到达方式/接驳")
    meal: str = Field(default="", description="餐饮建议，可空")
    note: str = Field(default="", description="备注：门票/预约/开放时间等，可空")


class DayPlan(BaseModel):
    day: int = Field(description="第几天")
    theme: str = Field(default="", description="当天主题")
    slots: List[Slot] = Field(default_factory=list, description="按时间排好的时段")
    est_cost: int = Field(default=0, description="当日预计花费(元, 整数)")


class DayFix(BaseModel):
    day: int = Field(description="需要调整的天")
    hint: str = Field(description="怎么调整的提示")


class Feasibility(BaseModel):
    ok: bool = Field(description="整体是否可行")
    issues: List[str] = Field(default_factory=list, description="发现的问题（给人看）")
    fixes: List[DayFix] = Field(default_factory=list, description="按天的修改建议")
