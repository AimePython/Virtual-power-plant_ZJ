from __future__ import annotations

import io
import os
from dataclasses import dataclass

import numpy as np
import pandas as pd
from flask import Flask, render_template_string, request, send_file, url_for

HOME_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>基线测算工具</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, "PingFang SC", sans-serif; margin: 28px auto; max-width: 1040px; line-height: 1.5; color: #1f2937; }
    h1 { margin: 0 0 6px; }
    .sub { color: #6b7280; margin-bottom: 18px; }
    .card { border: 1px solid #e5e7eb; border-radius: 12px; padding: 16px; margin-bottom: 14px; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    label { display: block; font-weight: 600; margin-bottom: 6px; }
    input, select, button { width: 100%; padding: 10px 12px; border: 1px solid #d1d5db; border-radius: 8px; font-size: 14px; box-sizing: border-box; }
    button { background: #2563eb; color: #fff; border: none; cursor: pointer; font-weight: 600; }
    button:hover { background: #1d4ed8; }
    .hint { font-size: 12px; color: #6b7280; margin-top: 6px; }
    .msg { padding: 10px 12px; border-radius: 8px; margin-top: 8px; }
    .ok { background: #ecfdf5; color: #065f46; border: 1px solid #a7f3d0; }
    .err { background: #fef2f2; color: #991b1b; border: 1px solid #fecaca; }
    @media (max-width: 720px) { .grid { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
  <h1>基线测算网页</h1>
  <div class="sub">上传数据后可下载结果，也可进入在线看板勾选单元格自动求和。</div>

  <div class="card">
    <h3>下载测算结果（Excel）</h3>
    <form method="post" action="{{ url_for('calculate') }}" enctype="multipart/form-data">
      <div class="grid">
        <div>
          <label>上传数据文件（.xlsx）</label>
          <input type="file" name="file" accept=".xlsx" required />
          <div class="hint">支持列：户号 / 户名 / DATA_DATE(或日期) / H00-H95</div>
        </div>
        <div>
          <label>响应日</label>
          <input type="date" name="target_date" required value="{{ default_date }}" />
          <div class="hint">工作日前5个工作日；周六/周日前3个同类日。</div>
        </div>
      </div>
      <div class="grid" style="margin-top:12px;">
        <div>
          <label>调节类型</label>
          <select name="response_type">
            <option value="positive">正调节</option>
            <option value="negative">负调节</option>
          </select>
        </div>
        <div>
          <label>输出粒度</label>
          <select name="output_granularity">
            <option value="48">48点（半小时）</option>
            <option value="96">96点（15分钟）</option>
            <option value="both">96点 + 48点</option>
          </select>
        </div>
      </div>
      <div style="margin-top: 14px;">
        <button type="submit">测算并下载</button>
      </div>
    </form>
  </div>

  <div class="card">
    <h3>在线看板（可勾选求和）</h3>
    <form method="post" action="{{ url_for('dashboard') }}" enctype="multipart/form-data">
      <div class="grid">
        <div>
          <label>上传数据文件（.xlsx）</label>
          <input type="file" name="file" accept=".xlsx" required />
        </div>
        <div>
          <label>响应日</label>
          <input type="date" name="target_date" required value="{{ default_date }}" />
        </div>
      </div>
      <div class="grid" style="margin-top:12px;">
        <div>
          <label>调节类型</label>
          <select name="response_type">
            <option value="positive">正调节</option>
            <option value="negative">负调节</option>
          </select>
        </div>
        <div>
          <label>看板粒度</label>
          <select name="board_granularity">
            <option value="48">48点（推荐）</option>
            <option value="96">96点</option>
          </select>
        </div>
      </div>
      <div style="margin-top: 14px;">
        <button type="submit">生成在线看板</button>
      </div>
    </form>
  </div>

  {% if message %}
    <div class="msg {{ 'ok' if ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

BOARD_HTML = """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <title>基线看板</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, "PingFang SC", sans-serif; margin: 20px; color: #111827; }
    .top { display:flex; justify-content:space-between; align-items:center; gap: 12px; flex-wrap: wrap; margin-bottom: 12px; }
    .btn { background:#2563eb; color:#fff; padding:8px 12px; border-radius:8px; text-decoration:none; font-weight:600; }
    .panel { border:1px solid #e5e7eb; border-radius:10px; padding:12px; margin-bottom:12px; }
    .small { color:#6b7280; font-size:12px; }
    .stats { display:flex; gap:12px; flex-wrap:wrap; }
    .stat { background:#f9fafb; border:1px solid #e5e7eb; border-radius:8px; padding:8px 12px; }
    .toolbar { display:grid; grid-template-columns:1fr 1.4fr 1fr; gap:12px; }
    .list { max-height:140px; overflow:auto; border:1px solid #e5e7eb; border-radius:8px; padding:8px; }
    .list label { display:block; font-size:13px; margin:4px 0; }
    .search { width:100%; box-sizing:border-box; margin:6px 0 8px; padding:7px 9px; border:1px solid #d1d5db; border-radius:8px; font-size:12px; }
    .table-wrap { overflow:auto; border:1px solid #e5e7eb; border-radius:8px; }
    table { border-collapse:collapse; font-size:12px; width:max-content; min-width:100%; }
    th, td { border:1px solid #e5e7eb; padding:4px 6px; text-align:right; white-space:nowrap; }
    thead th { position:sticky; top:0; background:#f3f4f6; z-index:3; }
    .time-col { position:sticky; left:0; z-index:2; background:#f9fafb; text-align:center; }
    .coef-col { position:sticky; right:110px; z-index:2; background:#f9fafb; text-align:center; }
    .sum-col { position:sticky; right:0; z-index:2; background:#f9fafb; text-align:right; font-weight:700; min-width: 96px; }
    .coef-input { width:72px; padding:3px 4px; border:1px solid #d1d5db; border-radius:6px; font-size:12px; text-align:right; }
    .operator-row th { top:0; background:#eef2ff; z-index:5; text-align:center; }
    .acct-row th { top:28px; background:#f3f4f6; z-index:4; text-align:center; }
    .num { cursor:pointer; }
    .selected { background:#fde68a !important; font-weight:700; }
    .row-check, .col-check { transform: scale(0.95); margin-right: 3px; }
    @media (max-width: 900px){ .toolbar { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <div class="top">
    <div>
      <h2 style="margin:0;">基线看板（{{ meta.granularity }}点）</h2>
      <div class="small">响应日：{{ meta.target_date }}｜调节：{{ meta.response_type }}｜已识别企业名称：{{ meta.operator_count }} 个</div>
    </div>
    <a class="btn" href="{{ url_for('home') }}">返回上传页</a>
  </div>

  <div class="panel stats">
    <div class="stat">已选单元格：<b id="selected-count">0</b></div>
    <div class="stat">已选求和(kW)：<b id="selected-sum">0.00</b></div>
    <div class="stat">当前显示户号：<b id="visible-count">0</b></div>
  </div>

  <div class="panel toolbar">
    <div>
      <div><b>企业名称筛选</b></div>
      <input id="enterprise-search" class="search" placeholder="搜索企业名称..." />
      <div class="list" id="operator-list"></div>
    </div>
    <div>
      <div><b>户号筛选</b></div>
      <input id="account-search" class="search" placeholder="搜索户号或户名..." />
      <div class="list" id="account-list"></div>
    </div>
    <div>
      <div><b>时段筛选</b></div>
      <input id="time-search" class="search" placeholder="搜索时刻，如 17:30" />
      <div class="list" id="time-list"></div>
    </div>
  </div>

  <div class="table-wrap">
    <table id="board-table"></table>
  </div>

  <script>
    const payload = {{ payload|tojson }};
    const operatorList = document.getElementById('operator-list');
    const accountList = document.getElementById('account-list');
    const timeList = document.getElementById('time-list');
    const enterpriseSearch = document.getElementById('enterprise-search');
    const accountSearch = document.getElementById('account-search');
    const timeSearch = document.getElementById('time-search');
    const table = document.getElementById('board-table');
    const selectedCountEl = document.getElementById('selected-count');
    const selectedSumEl = document.getElementById('selected-sum');
    const visibleCountEl = document.getElementById('visible-count');

    const selectedCells = new Set();
    const state = {
      operators: new Set(payload.operators),
      accounts: new Set(payload.accounts.map(a => a.key)),
      times: new Set(payload.times),
    };
    const bidCoefByTime = {};
    payload.times.forEach((_, i) => { bidCoefByTime[i] = 100; });

    function cellId(ti, ai){ return `${ti}|${ai}`; }
    function fmt(v){ return Number(v).toFixed(2); }

    function renderOperatorFilter(){
      operatorList.innerHTML = '';
      const kw = (enterpriseSearch.value || '').trim().toLowerCase();
      payload.operators
        .filter(op => !kw || op.toLowerCase().includes(kw))
        .forEach(op => {
        const id = `op-${op}`;
        const wrap = document.createElement('label');
        const checked = state.operators.has(op) ? 'checked' : '';
        wrap.innerHTML = `<input type="checkbox" ${checked} data-op="${op}" /> ${op}`;
        operatorList.appendChild(wrap);
      });
      operatorList.querySelectorAll('input[type=checkbox]').forEach(el => {
        el.addEventListener('change', e => {
          const op = e.target.dataset.op;
          if(e.target.checked){ state.operators.add(op); } else { state.operators.delete(op); }
          const allowed = new Set(payload.accounts.filter(a => state.operators.has(a.operator)).map(a => a.key));
          state.accounts = new Set([...state.accounts].filter(k => allowed.has(k)));
          if(state.accounts.size === 0){ allowed.forEach(k => state.accounts.add(k)); }
          renderAccountFilter();
          renderTable();
        });
      });
    }

    function renderAccountFilter(){
      accountList.innerHTML = '';
      const kw = (accountSearch.value || '').trim().toLowerCase();
      const visibleAccounts = payload.accounts.filter(a => {
        if(!state.operators.has(a.operator)) return false;
        if(!kw) return true;
        return String(a.key).toLowerCase().includes(kw) || String(a.name).toLowerCase().includes(kw);
      });
      visibleAccounts.forEach(a => {
        const checked = state.accounts.has(a.key) ? 'checked' : '';
        const wrap = document.createElement('label');
        wrap.innerHTML = `<input type="checkbox" ${checked} data-acct="${a.key}" /> ${a.key}（${a.name}）`;
        accountList.appendChild(wrap);
      });
      accountList.querySelectorAll('input[type=checkbox]').forEach(el => {
        el.addEventListener('change', e => {
          const key = e.target.dataset.acct;
          if(e.target.checked){ state.accounts.add(key); } else { state.accounts.delete(key); }
          renderTable();
        });
      });
    }

    function renderTimeFilter(){
      timeList.innerHTML = '';
      const kw = (timeSearch.value || '').trim();
      payload.times
        .filter(t => !kw || t.includes(kw))
        .forEach(t => {
          const checked = state.times.has(t) ? 'checked' : '';
          const wrap = document.createElement('label');
          wrap.innerHTML = `<input type="checkbox" ${checked} data-time="${t}" /> ${t}`;
          timeList.appendChild(wrap);
        });
      timeList.querySelectorAll('input[type=checkbox]').forEach(el => {
        el.addEventListener('change', e => {
          const t = e.target.dataset.time;
          if(e.target.checked){ state.times.add(t); } else { state.times.delete(t); }
          renderTable();
        });
      });
    }

    function recalcSelected() {
      let sum = 0;
      let count = 0;
      selectedCells.forEach(id => {
        const [ti, ai] = id.split('|').map(Number);
        if(payload.matrix[ti] && payload.matrix[ti][ai] !== undefined){
          sum += Number(payload.matrix[ti][ai]);
          count += 1;
        }
      });
      selectedCountEl.textContent = count;
      selectedSumEl.textContent = fmt(sum);
    }

    function toggleCell(td, ti, ai){
      const id = cellId(ti, ai);
      if(selectedCells.has(id)){
        selectedCells.delete(id);
        td.classList.remove('selected');
      }else{
        selectedCells.add(id);
        td.classList.add('selected');
      }
      recalcSelected();
    }

    function renderTable(){
      const acctIdx = payload.accounts
        .map((a, i) => ({...a, i}))
        .filter(a => state.operators.has(a.operator) && state.accounts.has(a.key));
      const visibleTimes = payload.times
        .map((t, i) => ({t, i}))
        .filter(x => state.times.has(x.t));
      visibleCountEl.textContent = acctIdx.length;

      let html = '';
      html += '<thead>';
      html += '<tr class="operator-row"><th class="time-col" rowspan="2">时刻</th>';
      acctIdx.forEach(a => { html += `<th>${a.operator}</th>`; });
      html += '<th class="coef-col" rowspan="2">申报系数(%)</th>';
      html += '<th class="sum-col" rowspan="2">申报值(MW)</th>';
      html += '</tr>';
      html += '<tr class="acct-row">';
      acctIdx.forEach(a => { html += `<th><input class="col-check" type="checkbox" data-col="${a.i}">${a.key}</th>`; });
      html += '</tr></thead><tbody>';

      visibleTimes.forEach(({t, i: ti}) => {
        html += `<tr><th class="time-col"><input class="row-check" type="checkbox" data-row="${ti}">${t}</th>`;
        let rowSum = 0;
        acctIdx.forEach(a => {
          const val = payload.matrix[ti][a.i];
          const id = cellId(ti, a.i);
          const cls = selectedCells.has(id) ? 'num selected' : 'num';
          rowSum += Number(val);
          html += `<td class="${cls}" data-ti="${ti}" data-ai="${a.i}">${fmt(val)}</td>`;
        });
        const coef = Number(bidCoefByTime[ti] ?? 100);
        // 申报系数是百分比，需要先除以100，再把kW转MW除以1000
        const declaredMw = (rowSum * (coef / 100.0)) / 1000.0;
        html += `<td class="coef-col"><input class="coef-input" type="number" step="0.01" min="0" data-row="${ti}" value="${coef}"></td>`;
        html += `<td class="sum-col" data-row-sum="${ti}">${fmt(declaredMw)}</td>`;
        html += '</tr>';
      });
      html += '</tbody>';
      table.innerHTML = html;

      table.querySelectorAll('td.num').forEach(td => {
        td.addEventListener('click', () => toggleCell(td, Number(td.dataset.ti), Number(td.dataset.ai)));
      });
      table.querySelectorAll('.row-check').forEach(chk => {
        chk.addEventListener('change', e => {
          const ti = Number(e.target.dataset.row);
          acctIdx.forEach(a => {
            const td = table.querySelector(`td[data-ti="${ti}"][data-ai="${a.i}"]`);
            const id = cellId(ti, a.i);
            if(e.target.checked && !selectedCells.has(id)){ selectedCells.add(id); td.classList.add('selected'); }
            if(!e.target.checked && selectedCells.has(id)){ selectedCells.delete(id); td.classList.remove('selected'); }
          });
          recalcSelected();
        });
      });
      table.querySelectorAll('.col-check').forEach(chk => {
        chk.addEventListener('change', e => {
          const ai = Number(e.target.dataset.col);
          visibleTimes.forEach(({i: ti}) => {
            const td = table.querySelector(`td[data-ti="${ti}"][data-ai="${ai}"]`);
            if(!td) return;
            const id = cellId(ti, ai);
            if(e.target.checked && !selectedCells.has(id)){ selectedCells.add(id); td.classList.add('selected'); }
            if(!e.target.checked && selectedCells.has(id)){ selectedCells.delete(id); td.classList.remove('selected'); }
          });
          recalcSelected();
        });
      });
      table.querySelectorAll('.coef-input').forEach(inp => {
        inp.addEventListener('change', e => {
          const ti = Number(e.target.dataset.row);
          const v = Number(e.target.value);
          bidCoefByTime[ti] = Number.isFinite(v) && v >= 0 ? v : 0;
          renderTable();
        });
      });
      recalcSelected();
    }

    enterpriseSearch.addEventListener('input', () => renderOperatorFilter());
    accountSearch.addEventListener('input', () => renderAccountFilter());
    timeSearch.addEventListener('input', () => renderTimeFilter());
    renderOperatorFilter();
    renderAccountFilter();
    renderTimeFilter();
    renderTable();
  </script>
</body>
</html>
"""

TIME_COLS = [f"H{i:02d}" for i in range(96)]


@dataclass
class RuleConfig:
    response_type: str
    calc_start_idx: int
    calc_end_idx: int
    threshold_desc: str
    threshold_func: callable


def get_rule_config(response_type: str) -> RuleConfig:
    if response_type == "negative":
        return RuleConfig("负调节", 0, 59, "Pavi > 1.25*Pav 视为异常并剔除", lambda pavi, pav: pavi > 1.25 * pav)
    return RuleConfig("正调节", 27, 91, "Pavi < 0.75*Pav 视为异常并剔除", lambda pavi, pav: pavi < 0.75 * pav)


def time_label_96(i: int) -> str:
    total = 15 + i * 15
    return "24:00" if total >= 1440 else f"{total // 60:02d}:{total % 60:02d}"


def time_label_48(i: int) -> str:
    total = 30 + i * 30
    return "24:00" if total >= 1440 else f"{total // 60:02d}:{total % 60:02d}"


def to_48_points(vals96: np.ndarray) -> np.ndarray:
    return np.array([(vals96[2 * i] + vals96[2 * i + 1]) / 2 for i in range(48)], dtype=float)


def normalize_input_df(df: pd.DataFrame) -> pd.DataFrame:
    col_map = {}
    for c in df.columns:
        s = str(c).strip()
        if s in ("DATA_DATE", "日期"):
            col_map[c] = "DATA_DATE"
        elif s == "户号":
            col_map[c] = "户号"
        elif s == "户名":
            col_map[c] = "户名"
        elif s in ("运营商名称", "虚拟电厂名称"):
            col_map[c] = "运营商名称"
    df = df.rename(columns=col_map)

    for c in ("户号", "户名", "DATA_DATE"):
        if c not in df.columns:
            raise ValueError(f"缺少必需列：{c}")
    if "运营商名称" not in df.columns:
        df["运营商名称"] = "未知运营商"

    for c in TIME_COLS:
        if c not in df.columns:
            raise ValueError(f"缺少时序列：{c}")
    df["DATA_DATE"] = pd.to_datetime(df["DATA_DATE"], errors="coerce")
    df = df.dropna(subset=["DATA_DATE"]).copy()
    df["运营商名称"] = df["运营商名称"].fillna("未知运营商").astype(str)
    for c in TIME_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def choose_ref_days(account_df: pd.DataFrame, target: pd.Timestamp, cfg: RuleConfig):
    w = target.weekday()
    if w < 5:
        required_n = 5
        candidates = [d for d in account_df["DATA_DATE"].unique() if pd.Timestamp(d) < target and pd.Timestamp(d).weekday() < 5]
    elif w == 5:
        required_n = 3
        candidates = [d for d in account_df["DATA_DATE"].unique() if pd.Timestamp(d) < target and pd.Timestamp(d).weekday() == 5]
    else:
        required_n = 3
        candidates = [d for d in account_df["DATA_DATE"].unique() if pd.Timestamp(d) < target and pd.Timestamp(d).weekday() == 6]
    candidates = [pd.Timestamp(d) for d in sorted(candidates, reverse=True) if (target - pd.Timestamp(d)).days <= 45]

    prepared = []
    for d in candidates:
        vals = account_df[account_df["DATA_DATE"] == d][TIME_COLS].astype(float).sum(axis=0).to_numpy(dtype=float)
        pavi = float(np.mean(vals[cfg.calc_start_idx : cfg.calc_end_idx + 1]))
        prepared.append({"date": d, "vals": vals, "pavi": pavi})

    selected = prepared[:required_n]
    initial = selected.copy()
    next_idx = required_n
    while selected:
        pav = float(np.mean([x["pavi"] for x in selected]))
        bad = [x for x in selected if cfg.threshold_func(x["pavi"], pav)]
        if not bad:
            threshold = (1.25 * pav) if cfg.response_type == "负调节" else (0.75 * pav)
            return initial, selected, threshold, required_n
        selected = [x for x in selected if not cfg.threshold_func(x["pavi"], pav)]
        while len(selected) < required_n and next_idx < len(prepared):
            selected.append(prepared[next_idx])
            next_idx += 1
        if len(selected) < required_n:
            threshold = (1.25 * pav) if cfg.response_type == "负调节" else (0.75 * pav)
            return initial, selected, threshold, required_n
    return initial, [], None, required_n


def compute_baseline_data(df: pd.DataFrame, target: pd.Timestamp, response_type: str):
    cfg = get_rule_config(response_type)
    labels96 = [time_label_96(i) for i in range(96)]
    labels48 = [time_label_48(i) for i in range(48)]

    summary_rows, ref_rows, errors, long_96, long_48 = [], [], [], [], []
    wide_96, wide_48 = {"时刻": labels96}, {"时刻": labels48}
    accounts_meta = []

    for acct, g in df.groupby("户号", sort=True):
        acct = int(acct)
        name = str(g["户名"].iloc[0])
        operator = str(g["运营商名称"].iloc[0]) if "运营商名称" in g.columns else "未知运营商"
        # 看板里“企业名称”按户号对应户名展示（与户号括号内名称一致）
        enterprise_name = name
        initial, refs, threshold, required_n = choose_ref_days(g, target, cfg)
        if len(refs) < 2:
            errors.append({"户号": acct, "户名": name, "运营商": operator, "错误": f"参考日不足，需{required_n}天，实际{len(refs)}天"})
            continue

        lowest = min(refs, key=lambda x: x["pavi"])
        final_refs = [x for x in refs if x["date"] != lowest["date"]]
        baseline96 = np.mean([x["vals"] for x in final_refs], axis=0)
        baseline48 = to_48_points(baseline96)
        key = str(acct)
        col = f"{key}_{name}"[:120]
        wide_96[col] = [round(float(v), 2) for v in baseline96]
        wide_48[col] = [round(float(v), 2) for v in baseline48]
        accounts_meta.append(
            {
                "key": key,
                "name": name,
                "operator": operator,
                "enterprise_name": enterprise_name,
                "column": col,
            }
        )

        summary_rows.append(
            {
                "运营商": operator,
                "户号": acct,
                "户名": name,
                "响应日": target.strftime("%Y-%m-%d"),
                "调节类型": cfg.response_type,
                "初始参考日": "、".join(x["date"].strftime("%Y-%m-%d") for x in initial),
                "初始参考日Pavi(kW)": "；".join(f"{x['date'].strftime('%Y-%m-%d')}={x['pavi']:.2f}" for x in initial),
                "阈值剔除日": "、".join(f"{x['date'].strftime('%Y-%m-%d')}({x['pavi']:.2f})" for x in initial if x["date"] not in {r["date"] for r in refs}) or "无",
                "递推后有效参考日": "、".join(x["date"].strftime("%Y-%m-%d") for x in refs),
                "剔除最低日": lowest["date"].strftime("%Y-%m-%d"),
                "最终参考日": "、".join(x["date"].strftime("%Y-%m-%d") for x in final_refs),
                "阈值(kW)": None if threshold is None else round(float(threshold), 2),
            }
        )
        for r in refs:
            ref_rows.append({"运营商": operator, "户号": acct, "户名": name, "参考日": r["date"].strftime("%Y-%m-%d"), "Pavi(kW)": round(float(r["pavi"]), 2)})
        for label, val in zip(labels96, baseline96):
            long_96.append({"运营商": operator, "户号": acct, "户名": name, "时刻": label, "基线负荷(kW)": round(float(val), 2)})
        for label, val in zip(labels48, baseline48):
            long_48.append({"运营商": operator, "户号": acct, "户名": name, "时刻": label, "基线负荷(kW)": round(float(val), 2)})

    return {
        "config": cfg,
        "summary_df": pd.DataFrame(summary_rows).sort_values(by=["运营商", "户号"]) if summary_rows else pd.DataFrame(),
        "ref_df": pd.DataFrame(ref_rows).sort_values(by=["运营商", "户号", "参考日"]) if ref_rows else pd.DataFrame(),
        "err_df": pd.DataFrame(errors),
        "wide96_df": pd.DataFrame(wide_96),
        "wide48_df": pd.DataFrame(wide_48),
        "long96_df": pd.DataFrame(long_96),
        "long48_df": pd.DataFrame(long_48),
        "accounts_meta": accounts_meta,
    }


def build_result_workbook(df: pd.DataFrame, target: pd.Timestamp, response_type: str, output_granularity: str) -> io.BytesIO:
    res = compute_baseline_data(df, target, response_type)
    cfg = res["config"]
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        pd.DataFrame([{
            "响应日": target.strftime("%Y-%m-%d"),
            "调节类型": cfg.response_type,
            "阈值规则": cfg.threshold_desc,
            "计算时段(H索引)": f"H{cfg.calc_start_idx:02d}-H{cfg.calc_end_idx:02d}",
            "H列映射": "H00=00:15，H01=00:30，...，H95=24:00",
            "48点转换": "第i点=(H[2i]+H[2i+1])/2，对应00:30、01:00、...、24:00",
        }]).to_excel(writer, sheet_name="计算口径", index=False)
        res["summary_df"].to_excel(writer, sheet_name="计算说明", index=False)
        res["ref_df"].to_excel(writer, sheet_name="参考日Pavi明细", index=False)
        if output_granularity in ("96", "both"):
            res["wide96_df"].to_excel(writer, sheet_name="基线96点_所有户号", index=False)
            res["long96_df"].to_excel(writer, sheet_name="基线96点_长表", index=False)
        if output_granularity in ("48", "both"):
            res["wide48_df"].to_excel(writer, sheet_name="基线48点_所有户号", index=False)
            res["long48_df"].to_excel(writer, sheet_name="基线48点_长表", index=False)
        if not res["err_df"].empty:
            res["err_df"].to_excel(writer, sheet_name="异常", index=False)
    bio.seek(0)
    return bio


def build_dashboard_payload(res: dict, target: pd.Timestamp, granularity: str):
    wide_df = res["wide48_df"] if granularity == "48" else res["wide96_df"]
    labels = wide_df["时刻"].tolist()
    columns = [c for c in wide_df.columns if c != "时刻"]
    meta_by_col = {m["column"]: m for m in res["accounts_meta"]}
    accounts = []
    for c in columns:
        m = meta_by_col.get(
            c,
            {
                "key": c.split("_")[0],
                "name": c.split("_", 1)[1] if "_" in c else c,
                "operator": "未知运营商",
                "enterprise_name": c.split("_", 1)[1] if "_" in c else c,
            },
        )
        accounts.append(
            {
                "key": str(m["key"]),
                "name": str(m["name"]),
                "operator": str(m["enterprise_name"]),
            }
        )
    matrix = wide_df[columns].to_numpy(dtype=float).tolist()
    operators = sorted({a["operator"] for a in accounts})
    return {
        "times": labels,
        "accounts": accounts,
        "operators": operators,
        "matrix": matrix,
        "meta": {
            "target_date": target.strftime("%Y-%m-%d"),
            "response_type": res["config"].response_type,
            "granularity": granularity,
            "operator_count": len(operators),
        },
    }


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


@app.get("/")
def home():
    return render_template_string(HOME_HTML, message=None, ok=True, default_date=pd.Timestamp.today().strftime("%Y-%m-%d"))


@app.post("/calculate")
def calculate():
    file = request.files.get("file")
    target_date = request.form.get("target_date", "").strip()
    response_type = request.form.get("response_type", "positive").strip()
    output_granularity = request.form.get("output_granularity", "48").strip()
    if not file or not file.filename:
        return render_template_string(HOME_HTML, message="请先上传 Excel 文件。", ok=False, default_date=target_date or pd.Timestamp.today().strftime("%Y-%m-%d"))
    try:
        target = pd.Timestamp(target_date)
        raw = pd.read_excel(file, sheet_name=0)
        norm = normalize_input_df(raw)
        output = build_result_workbook(norm, target, response_type, output_granularity)
        fname = f"基线测算结果_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(output, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True, download_name=fname)
    except Exception as exc:
        return render_template_string(HOME_HTML, message=f"测算失败：{exc}", ok=False, default_date=target_date or pd.Timestamp.today().strftime("%Y-%m-%d"))


@app.post("/dashboard")
def dashboard():
    file = request.files.get("file")
    target_date = request.form.get("target_date", "").strip()
    response_type = request.form.get("response_type", "positive").strip()
    granularity = request.form.get("board_granularity", "48").strip()
    if not file or not file.filename:
        return render_template_string(HOME_HTML, message="请先上传 Excel 文件。", ok=False, default_date=target_date or pd.Timestamp.today().strftime("%Y-%m-%d"))
    try:
        target = pd.Timestamp(target_date)
        raw = pd.read_excel(file, sheet_name=0)
        norm = normalize_input_df(raw)
        res = compute_baseline_data(norm, target, response_type)
        payload = build_dashboard_payload(res, target, granularity if granularity in ("48", "96") else "48")
        return render_template_string(BOARD_HTML, payload=payload, meta=payload["meta"])
    except Exception as exc:
        return render_template_string(HOME_HTML, message=f"生成看板失败：{exc}", ok=False, default_date=target_date or pd.Timestamp.today().strftime("%Y-%m-%d"))


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8501"))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
