"""
Self-check for the Chapter Update backend.

Builds throwaway workbooks shaped like the real ones, rolls a chapter forward,
and asserts the things that were actually wrong in production before this logic
was fixed:

  * new rows land directly AFTER the previous chapter's block, not appended at
    the bottom of the sheet
  * all three number forms move together (map id / Chinese numerals / display)
  * every chest series (60/70/80) gets ALL its rows copied, not just the last
  * the harvest atlas follows fruit_id // 10, not the previous row's value
  * running twice is refused rather than silently duplicating the chapter

    python test_chapter_core.py
"""
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl

import chapter_core
from chapter_core import Cfg, _cn_num

# Check names carry ★ and Chinese sheet names; a piped Windows console is
# cp1252 and would die on them.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

FAILS = []


def check(ok, msg, detail=""):
    if ok:
        print("OK   " + msg)
    else:
        print("FAIL " + msg + (("\n     " + detail) if detail else ""))
        FAILS.append(msg)


def _sheet(wb, name, header, rows):
    ws = wb.create_sheet(name) if name not in wb.sheetnames else wb[name]
    ws.append(header)
    for r in rows:
        ws.append(r)
    return ws


def build_levels(design: Path, chapters=(78, 79)):
    """pve_level / pve_layout shaped like the real ones.

    Column positions match production exactly, because QA locates G/M/O/P/Q by
    their Chinese headers and the U-AB consistency window by offset.
    """
    hdr = ["关卡ID", "Alias", "Stock布局配置", "胜利条件", "初始Building1",
           "初始Building2", "限制回合数", "控制触发条件", "控制类型",
           "偏移控制触发预设方案", "偏移控制触发条件", "偏移量调控值",
           "动态方案", "连击奖励循环列表", "困难程度\n0-普通 1-困难 2-噩梦",
           "星分数条件", "星最大分数", "星基础奖励", "星基础奖励$E1",
           "星基础奖励$E3", "星基础奖励$E4-E6", "星基础奖励$J70",
           "额外分数转化金币系数", "额外分数转化金币系数$E4-E6", "门票数",
           "门票数$E1-E3", "门票数$E4-E6", "门票数$J70"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "level"
    ws.append(hdr)
    for ch in chapters:
        for n in range(1, 201):
            star3 = 30 + n
            row = [None] * 28
            row[0] = ch * 1000 + n
            row[1] = f"章节{ch}关卡{n}"
            row[2] = f"章节{ch}关卡{n}"
            # Two levels per chapter carry a collect-card objective.
            row[3] = "收集卡:1" if n in (5, 9) else "清空Stocks:0"
            row[6] = 4 + (n % 4)                      # hands, always 4-7
            row[12] = "中等" if n % 2 else None        # difficulty text
            row[14] = (1 if n == 20 else 2 if n == 40 else 0)
            row[15] = f"10,20,{star3}"                 # 1★/2★/3★
            row[16] = round(1.2 * star3)               # = 1.2 x 3★
            for c in range(17, 28):                    # U..AB constant per level
                row[c] = f"v{n}"
            ws.append(row)
    wb.save(design / "pve_level.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "layout"
    ws.append(["关卡ID", "Alias", "布局"])
    for ch in chapters:
        for n in range(1, 201):
            layout = '{"card":14,}' if n in (5, 9) else '{"card":3,}'
            ws.append([ch * 1000 + n, f"章节{ch}关卡{n}", layout])
    wb.save(design / "pve_layout.xlsx")


def build_design(design: Path):
    """Minimal but structurally faithful stand-ins for the real workbooks."""
    design.mkdir(parents=True, exist_ok=True)

    # ---- pve_chapter.xlsx: chapter / chest_plan / map / reward_zone / harvest
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _sheet(wb, "chapter",
           ["章节ID", "Alias", "章节数字", "地图列表", "奖励区域列表", "章节局内背景",
            "音效", "解锁图片", "SP1", "SP2", "SP3", "Tip", "道具", "题图", "入口icon",
            "礼包", "title"],
           [[77000, "第七十七章节", 77, "第七十七章节地图1,第七十七章节地图13", "1305,1308",
             "章节局内背景图77", "PVE游戏内背景音效2", "章节77解锁图片",
             "高三星低让利第77章StarPass方案", "高三星更低让利第77章StarPass方案",
             "高三星免费第77章StarPass方案", "5,14,20", "章节77通行证",
             "StarPass界面题图77", "StarPass入口Icon77", "章节77解锁礼包", "地图跳转title77"],
            [78000, "第七十八章节", 78, "第七十八章节地图1,第七十八章节地图13", "1309,1312",
             "章节局内背景图78", "PVE游戏内背景音效3", "章节78解锁图片",
             "高三星低让利第78章StarPass方案", "高三星更低让利第78章StarPass方案",
             "高三星免费第78章StarPass方案", "5,14,20", "章节78通行证",
             "StarPass界面题图78", "StarPass入口Icon78", "章节78解锁礼包", "地图跳转title78"],
            # A trailing non-chapter row: new rows must be inserted ABOVE it.
            ["//END", None, None, None, None, None, None, None, None, None, None,
             None, None, None, None, None, None]])
    _sheet(wb, "chest_plan", ["方案ID", "Alias", "星数上限", "赠送比率", "星级奖励列表"],
           [["6077000", "高三星低让利第77章StarPass方案", "600", None, "高三星低让利章节77星0,"],
            ["6078000", "高三星低让利第78章StarPass方案", "600", None, "高三星低让利章节78星0,"],
            ["7077000", "高三星更低让利第77章StarPass方案", "600", None, "高三星更低让利章节77星0,"],
            ["7078000", "高三星更低让利第78章StarPass方案", "600", None, "高三星更低让利章节78星0,"],
            ["8077000", "高三星免费第77章StarPass方案", "600", None, "高三星免费章节77星0,"],
            ["8078000", "高三星免费第78章StarPass方案", "600", None, "高三星免费章节78星0,"]])
    map_rows = []
    for i in range(1, 14):
        # Real data: row NN001 carries direction 0 (down) and NN012 carries the
        # Tip (2) while it is the last chapter in the game.
        dir_pos, dir_type = (None, None)
        if i == 1:
            dir_pos, dir_type = "33,-751", 0
        elif i == 12:
            dir_pos, dir_type = "-204,831", 2
        map_rows.append([78000 + i, f"第七十八章节地图{i}", f"场景七十八地图片{i}",
                         f"章节78关卡{i}", "0,0", "327", "1,1", "1.5,1.5",
                         dir_pos, dir_type])
    _sheet(wb, "map",
           ["地图ID", "Alias", "地图PIC", "关卡列表", "关卡坐标", "采集物列表",
            "采集物坐标", "采集物旋转", "方向键位置", "方向键类型"], map_rows)
    _sheet(wb, "reward_zone",
           ["奖励区域ID", "起始关卡ID", "结束关卡ID", "奖励加成", "采集物种类列表", "加成3", "加成3b"],
           [[1309, "章节78关卡1", "章节78关卡50", 0, 327, "0", "0"],
            [1310, "章节78关卡51", "章节78关卡100", 0, 328, "0", "0"],
            [1311, "章节78关卡101", "章节78关卡150", 0, 329, "0", "0"],
            [1312, "章节78关卡151", "章节78关卡200", 0, 330, "0", "0"]])
    _sheet(wb, "harvest",
           ["采集物ID", "未解锁", "等待收获", "可收获", "类型", "所在图集"],
           [[327, "FRUIT327_1", "FRUIT327_2", "FRUIT327_3", 1, "MAP_FRUIT_32"],
            [328, "FRUIT328_1", "FRUIT328_2", "FRUIT328_3", 1, "MAP_FRUIT_32"],
            [329, "FRUIT329_1", "FRUIT329_2", "FRUIT329_3", 1, "MAP_FRUIT_32"],
            [330, "FRUIT330_1", "FRUIT330_2", "FRUIT330_3", 1, "MAP_FRUIT_33"]])
    wb.save(design / "pve_chapter.xlsx")

    # ---- asset_ref.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "key_asset"
    ws.append(["资源名", "Alias", "资源路径"])
    for i in range(1, 14):
        ws.append([f"MAP_78_{i}", f"场景七十八地图片{i}",
                   f"Assets/Map/dynamic_map_78000/bg_78000/map_78_{i}.jpg"])
    ws.append(["PVE_BG_78", "章节局内背景图78",
               "Assets/Map/dynamic_map_78000/pve_78000/pve_bg_78.jpg"])
    ws.append(["CHAPTER_UNLOCK_78", "章节78解锁图片",
               "Assets/ResourcesAb/Texture/ChapeDoneBg/img_chapter_unlock_78.png"])
    for a in (32, 33):
        ws.append([f"MAP_FRUIT_{a}", f"MAP_FRUIT_{a}", f"Assets/ResourcesAb/Icon/MapIcon{a}"])
    ws.append(["STARPASS_CHAPTER_BG_78", "StarPass界面题图78",
               "Assets/Map/dynamic_map_78000/starpass_78000/img_chapterbg78.png"])
    ws.append(["STARPASS_ENTRY_ICON_78", "StarPass入口Icon78",
               "Assets/Map/dynamic_map_78000/starpass_78000/icon_cheststart78.png"])
    ws.append(["CHAPTER_TITLE_BG_78", "地图跳转title78",
               "Assets/ResourcesAb/Texture/AdventureMapBg/img_chapter78.png"])
    ws.append(["ZZZ_SENTINEL", "must stay last", ""])
    wb.save(design / "asset_ref.xlsx")

    # ---- icon.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "icon_chapter"
    ws.append(["icon名", "Alias", "精灵名字", "图集", None])
    for state in (1, 2, 3):
        for fid in (327, 328, 329, 330):
            ws.append([f"FRUIT{fid}_{state}", f"场景七十八收获物{fid}状态{state}",
                       f"fruit{fid}_{state}", None, None])
    wb.save(design / "icon.xlsx")

    # ---- chest.xlsx (three series, 2 rows each -> all 6 must be copied)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet in ("chest_conditions", "chest_rewards"):
        rows = []
        for pfx in ("60", "70", "80"):
            for n in (1, 2):
                if sheet == "chest_conditions":
                    # A, Alias, 收集数, 章节ID(map id), 免费宝箱
                    rows.append([int(f"{pfx}78{n:03d}"), f"高三星章节78星{n}",
                                 str(n), 78000, f"高三星章节78_免费{n}"])
                else:
                    # chest_rewards: col D is reward CONTENT (must not change),
                    # col E is the bare display chapter number (must change).
                    rows.append([int(f"{pfx}78{n:03d}"), f"高三星章节78_免费{n}",
                                 "展示道具", "1000,200", 78])
        _sheet(wb, sheet, ["ID", "Alias", "desc", "D", "colE"], rows)
    wb.save(design / "chest.xlsx")

    # ---- item.xlsx / store.xlsx / notification.xlsx
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "item"
    ws.append(["id", "name"])
    ws.append([4076, "章节77通行证"])
    ws.append([4077, "章节78通行证"])
    ws.append(["//头像框", "not an item"])
    wb.save(design / "item.xlsx")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sheet in ("store_recharge", "store_package"):
        _sheet(wb, sheet, ["id", "name"],
               [[40077, "章节77解锁礼包"], [40078, "章节78解锁礼包"], ["//end", "tail"]])
    wb.save(design / "store.xlsx")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "server_notification"
    ws.append(["id", "name", "desc", "flag", "e", "f", "g", "push", "i"])
    # Real rows carry a 3-day window (start .. start+2); QA checks that span,
    # and the reward step advances it by the chapter cadence.
    ws.append([501, "78章节上线", "新关卡章节开放 ", True,
               "2026/09/04 00:00:00", "2026/09/06 00:00:00", None,
               "PUSH_NOTE_NEW_CHAPTER2", None])
    wb.save(design / "notification.xlsx")

    build_levels(design)


def col(ws, idx):
    return [r[idx] for r in ws.iter_rows(min_row=2, values_only=True)]


def main():
    tmp = Path(tempfile.mkdtemp(prefix="chapter_core_test_"))
    try:
        branch = tmp / "trunk"
        design = branch / "design" / "DesignData" / "WS" / "design"
        build_design(design)

        # Art assets: a source map folder + the two texture placeholders.
        cmap = branch / "develop" / "client" / "Skipbo" / "Assets" / "Map"
        (cmap / "dynamic_map_78000" / "bg_78000").mkdir(parents=True)
        (cmap / "dynamic_map_78000" / "bg_78000" / "map_78_1.jpg").write_text("x")
        (cmap / "dynamic_map_78000" / "bg_78000" / "map_78_1.jpg.meta").write_text("m")
        tex = branch / "develop" / "client" / "Skipbo" / "Assets" / "ResourcesAb" / "Texture"
        (tex / "AdventureMapBg").mkdir(parents=True)
        (tex / "AdventureMapBg" / "img_chapter68.png").write_text("p")
        (tex / "ChapeDoneBg").mkdir(parents=True)
        (tex / "ChapeDoneBg" / "img_chapter_unlock_68.png").write_text("p")

        # display_offset 10 -> internal 78 shows as 章节68? No: the fixture uses
        # display == internal (offset 0) so the Chinese numerals stay 七十八.
        cfg = Cfg(branch_root=branch, workspace="WS", old_ch=78, display_offset=0)
        quiet = lambda *a, **k: None

        check(_cn_num(78) == "七十八" and _cn_num(79) == "七十九" and _cn_num(101) == "一百零一",
              "Chinese numeral conversion (78/79/101)")

        chapter_core.run_step4_reward_config(cfg, quiet)
        chapter_core.run_step6_asset_config(cfg, quiet)

        # ---- chest: every row of every series copied, and in place ----------
        wb = openpyxl.load_workbook(design / "chest.xlsx")
        ids = [str(v) for v in col(wb["chest_conditions"], 0)]
        check(len([i for i in ids if i.startswith("6079")]) == 2
              and len([i for i in ids if i.startswith("7079")]) == 2
              and len([i for i in ids if i.startswith("8079")]) == 2,
              "chest: all 3 series copied ALL their rows (2 each), not just the last",
              "got: " + ",".join(ids))
        check(ids.index("6079001") == ids.index("6078002") + 1,
              "chest: new 60-series block sits directly after the old one",
              "got: " + ",".join(ids))

        # ---- item / store: inserted after the range, before the tail -------
        wb = openpyxl.load_workbook(design / "item.xlsx")
        item_ids = [str(v) for v in col(wb["item"], 0)]
        check(item_ids == ["4076", "4077", "4078", "//头像框"],
              "item: 4078 inserted after 4077 and before the trailing marker",
              "got: " + ",".join(item_ids))

        wb = openpyxl.load_workbook(design / "store.xlsx")
        store_ids = [str(v) for v in col(wb["store_recharge"], 0)]
        check(store_ids == ["40077", "40078", "40079", "//end"],
              "store: 40079 inserted in place, not appended after the tail",
              "got: " + ",".join(store_ids))

        # ---- chapter sheet: every number form moved -------------------------
        wb = openpyxl.load_workbook(design / "pve_chapter.xlsx")
        ws = wb["chapter"]
        new = [r for r in ws.iter_rows(min_row=2, max_col=17, values_only=True)
               if r[0] == 79000]
        check(len(new) == 1, "chapter: exactly one new row")
        r = new[0]
        check(r[1] == "第七十九章节", "chapter col B Chinese numeral rolled", repr(r[1]))
        check(r[2] == 79, "chapter col C display number rolled", repr(r[2]))
        check(r[3] == "第七十九章节地图1,第七十九章节地图13", "chapter col D map list rolled", repr(r[3]))
        check(r[4] == "1313,1316", "chapter col E reward zones advanced by 4", repr(r[4]))
        check(r[5] == "章节局内背景图79", "chapter col F background rolled", repr(r[5]))
        check(r[13] == "StarPass界面题图79" and r[14] == "StarPass入口Icon79"
              and r[16] == "地图跳转title79",
              "chapter cols N/O/Q rolled (the ones the old code silently skipped)",
              f"{r[13]!r} {r[14]!r} {r[16]!r}")
        ch_ids = [str(v) for v in col(ws, 0)]
        check(ch_ids == ["77000", "78000", "79000", "//END"],
              "chapter: new row inserted after 78000, above the trailing marker",
              "got: " + ",".join(ch_ids))

        # ---- chest_plan: all THREE variants ---------------------------------
        plan_ids = [str(v) for v in col(wb["chest_plan"], 0)]
        check(sorted(i for i in plan_ids if i.endswith("79000")) == ["6079000", "7079000", "8079000"],
              "chest_plan: all 3 StarPass variants created", "got: " + ",".join(plan_ids))

        # ---- map: 13 rows, ids +1000, Chinese numerals rolled ---------------
        map_new = [r for r in wb["map"].iter_rows(min_row=2, max_col=4, values_only=True)
                   if isinstance(r[0], int) and 79000 < r[0] < 80000]
        check(len(map_new) == 13, f"map: 13 new rows (got {len(map_new)})")
        check(map_new[0][1] == "第七十九章节地图1" and map_new[0][2] == "场景七十九地图片1",
              "map: cols B/C Chinese numerals rolled", repr(map_new[0][1:3]))

        # ---- harvest: atlas is fruit_id // 10, NOT copied from the old row --
        hv = {r[0]: r[5] for r in wb["harvest"].iter_rows(min_row=2, max_col=6, values_only=True)}
        check([hv.get(f) for f in (331, 332, 333, 334)] == ["MAP_FRUIT_33"] * 4,
              "harvest: fruits 331-334 all map to MAP_FRUIT_33 (old code said _32)",
              repr({f: hv.get(f) for f in (331, 332, 333, 334)}))

        # ---- reward_zone: ids + fruit refs advanced -------------------------
        rz = [(r[0], r[1], r[4]) for r in wb["reward_zone"].iter_rows(min_row=2, max_col=5, values_only=True)]
        new_rz = [x for x in rz if x[0] in (1313, 1314, 1315, 1316)]
        check(len(new_rz) == 4 and new_rz[0][1] == "章节79关卡1" and new_rz[0][2] == 331,
              "reward_zone: 4 zones added, level text + fruit ref advanced", repr(new_rz[:1]))

        # ---- icon_chapter: 12 rows, new fruit ids + Chinese numerals --------
        wb = openpyxl.load_workbook(design / "icon.xlsx")
        icon_new = [r for r in wb["icon_chapter"].iter_rows(min_row=2, max_col=3, values_only=True)
                    if r[1] and "场景七十九" in str(r[1])]
        check(len(icon_new) == 12, f"icon_chapter: 12 new rows (got {len(icon_new)})")
        check(all(str(r[0]).startswith(("FRUIT331", "FRUIT332", "FRUIT333", "FRUIT334"))
                  for r in icon_new),
              "icon_chapter: fruit ids advanced (old code cloned 327-330 verbatim)")

        # ---- asset_ref: all 6 display groups, each in place -----------------
        wb = openpyxl.load_workbook(design / "asset_ref.xlsx")
        names = [str(v) for v in col(wb["key_asset"], 0)]
        for grp in ("PVE_BG_79", "CHAPTER_UNLOCK_79", "STARPASS_CHAPTER_BG_79",
                    "STARPASS_ENTRY_ICON_79", "CHAPTER_TITLE_BG_79"):
            check(grp in names, f"asset_ref: {grp} created")
        check(names.count("MAP_79_13") == 1 and names.index("MAP_79_1") == names.index("MAP_78_13") + 1,
              "asset_ref: MAP_79 block sits directly after MAP_78_13")
        check(names[-1] == "ZZZ_SENTINEL",
              "asset_ref: nothing appended past the end of the sheet", "last=" + names[-1])
        paths = {r[0]: r[2] for r in wb["key_asset"].iter_rows(min_row=2, max_col=3, values_only=True)}
        check(paths.get("PVE_BG_79") == "Assets/Map/dynamic_map_79000/pve_79000/pve_bg_79.jpg",
              "asset_ref: path carries BOTH the new map id and the new display number",
              repr(paths.get("PVE_BG_79")))
        aliases = {r[0]: r[1] for r in wb["key_asset"].iter_rows(min_row=2, max_col=3, values_only=True)}
        check(aliases.get("MAP_79_1") == "场景七十九地图片1",
              "asset_ref: MAP alias uses Chinese numerals", repr(aliases.get("MAP_79_1")))
        check("MAP_FRUIT_34" not in names,
              "asset_ref: no phantom MAP_FRUIT_34 (fruits 331-334 fit the existing _33)")

        # ---- doc rules that the tool used to miss ---------------------------
        # "Sheet chest_reward has a different format it only has chapter
        # number, simply change it to new chapter number" -- a bare 78 with no
        # 章节 prefix, which the context-requiring replacer cannot see.
        wb = openpyxl.load_workbook(design / "chest.xlsx")
        wsr = wb["chest_rewards"]
        new_rew = [[c.value for c in r] for r in wsr.iter_rows(min_row=2)
                   if str(r[0].value or "").startswith(("6079", "7079", "8079"))]
        check(len(new_rew) == 6, f"chest_rewards: 6 new rows (got {len(new_rew)})")
        check(all(r[4] == 79 for r in new_rew),
              "chest_rewards col E (bare chapter number) advanced to 79",
              repr([r[4] for r in new_rew]))
        check(all(r[3] == "1000,200" for r in new_rew),
              "chest_rewards col D (reward content) left alone, per the doc",
              repr([r[3] for r in new_rew]))
        wsc = wb["chest_conditions"]
        new_cond = [[c.value for c in r] for r in wsc.iter_rows(min_row=2)
                    if str(r[0].value or "").startswith("6079")]
        check(all(r[3] == 79000 for r in new_cond),
              "chest_conditions col D (map id) advanced to 79000",
              repr([r[3] for r in new_cond]))

        # "In column J replace the number from 1 to 2 for latest chapter and
        # 2 to 1 for previous chapter" -- the Tip on the last chapter becomes
        # an up-arrow once a newer chapter exists.
        wb = openpyxl.load_workbook(design / "pve_chapter.xlsx")
        wsm = wb["map"]
        jvals = {r[0].value: r[9].value for r in wsm.iter_rows(min_row=2, max_col=10)
                 if isinstance(r[0].value, int)}
        check(jvals.get(79012) == 2,
              f"map: new chapter row 79012 carries the Tip (2), got {jvals.get(79012)}")
        check(jvals.get(78012) == 1,
              f"map: previous chapter row 78012 demoted 2 -> 1, got {jvals.get(78012)}")
        check(jvals.get(79001) == 0 and jvals.get(78001) == 0,
              "map: the down-arrow rows (NN001) are untouched")

        # ---- re-running is refused, not silently duplicated -----------------
        try:
            chapter_core.run_step6_asset_config(cfg, quiet)
            check(False, "asset config refuses a second run", "it ran again instead")
        except RuntimeError as e:
            check("already contains chapter" in str(e),
                  "asset config refuses a second run with a revert hint", str(e)[:80])

        # ---- art assets: copy, rename, strip .meta --------------------------
        chapter_core.run_step5_art_assets(cfg, quiet)
        dst = cmap / "dynamic_map_79000"
        check(dst.is_dir(), "art: dynamic_map_79000 created")
        check((dst / "bg_79000" if (dst / "bg_79000").exists() else dst / "bg_78000").exists(),
              "art: inner folder present")
        check(not list(dst.rglob("*.meta")), "art: stale .meta files removed")
        check(any(p.name == "map_79_1.jpg" for p in dst.rglob("*")),
              "art: files renamed to the new chapter",
              repr([p.name for p in dst.rglob("*") if p.is_file()]))
        # ---- toolchain step wiring -------------------------------------------
        # These steps shell out to SVN / Unity / the simulator, so they are not
        # executed here. What IS tested is everything that decides WHERE they
        # look and whether they fail loudly: a wrong tools/ path is the most
        # likely way this breaks on someone else's checkout.
        repo = tmp                      # branch is <tmp>/trunk, so tools is <tmp>/tools
        (repo / "tools" / "Level_Builder3" / "Magic_Builder" / "Assets" / "Py").mkdir(parents=True)
        (repo / "tools" / "Map_Builder" / "Magic_Builder").mkdir(parents=True)
        guessed = Cfg(branch_root=branch, workspace="WS", old_ch=78, display_offset=0)
        check(guessed.lvl_py == repo / "tools" / "Level_Builder3" / "Magic_Builder" / "Assets" / "Py",
              "tools guess: Level Builder Py resolves next to the checkout",
              str(guessed.lvl_py))
        check(guessed.map_builder == repo / "tools" / "Map_Builder" / "Magic_Builder",
              "tools guess: Map Builder project resolves next to the checkout",
              str(guessed.map_builder))
        check(guessed.map_scenes == guessed.map_builder / "Assets" / "Scenes"
              and guessed.map_py == guessed.map_builder / "Assets" / "Py"
              and guessed.map_output == guessed.map_builder / "__Output",
              "Map Builder scenes/py/output all hang off the project folder")

        # A branch two levels down (branches/Misc) must still find <repo>/tools —
        # this is exactly why these paths cannot be derived from the branch alone.
        deep = repo / "branches" / "Misc"
        (deep / "design").mkdir(parents=True, exist_ok=True)
        deep_cfg = Cfg(branch_root=deep, workspace="WS", old_ch=78, display_offset=0)
        check(deep_cfg.map_builder == repo / "tools" / "Map_Builder" / "Magic_Builder",
              "a branches/<name> checkout still finds <repo>/tools",
              str(deep_cfg.map_builder))

        override = Cfg(branch_root=branch, workspace="WS", old_ch=78, display_offset=0,
                       lvl_py=repo / "elsewhere" / "py",
                       map_builder=repo / "elsewhere" / "mb")
        check(override.lvl_py == repo / "elsewhere" / "py"
              and override.map_scenes == repo / "elsewhere" / "mb" / "Assets" / "Scenes",
              "explicit tool paths win over the guess")

        check(guessed.design_root() == branch / "design" / "DesignData" / "WS",
              "design_root() is where coder_convert.bat is looked for",
              str(guessed.design_root()))

        # Each toolchain step must fail with a path in the message, not a
        # bare KeyError/TypeError, when its prerequisite is absent.
        for label, fn, needle in [
            ("Magic Builder", chapter_core.run_step1_magic_builder, "CondFilterConf.json"),
            ("Screenshot Upload", chapter_core.run_step2_upload, "Level Builder Py folder"),
            ("Import Config", chapter_core.run_step8_import_config, "coder_convert.bat"),
            ("Simulation", chapter_core.run_step9_simulation, "run_conf_2.json"),
        ]:
            try:
                fn(guessed, quiet)
                ok, detail = False, "did not raise"
            except Exception as e:
                ok = needle in str(e) or needle in repr(e)
                detail = f"{type(e).__name__}: {str(e)[:70]}"
            check(ok, f"{label} fails with a usable message when its files are missing",
                  detail)

        check(chapter_core._NO_PYTHON_MSG.startswith("No system Python"),
              "the no-Python message is intact (it is what users see for step 1)")
        # ---- QA validation ---------------------------------------------------
        # The whole point of QA is to be trusted when it says PASS, so the
        # fixture is built to satisfy every automated rule and the run must come
        # back with zero FAILs. Each rule is then broken on purpose below to
        # prove the check actually looks at the data.
        qa_path = chapter_core.run_qa_validation(cfg, quiet)
        wb = openpyxl.load_workbook(qa_path)
        qa = [[c.value for c in r] for r in wb["QA Validation"].iter_rows(min_row=2)]
        fails = [r for r in qa if r[2] == "FAIL"]
        check(not fails, "QA: a correctly built chapter produces zero FAILs",
              "; ".join(f"{r[1]} ({r[3]})" for r in fails[:4]))
        check(any(r[2] == "MANUAL" for r in qa),
              "QA: human-only checklist items are carried into the report")
        check(len(qa) >= 40, f"QA: report covers the checklist ({len(qa)} rows)")

        # Break one rule from each family and confirm QA notices.
        def qa_rows(mutate):
            import copy as _copy
            probe = tmp / "probe"
            if probe.exists():
                shutil.rmtree(probe)
            shutil.copytree(branch, probe)
            pdesign = probe / "design" / "DesignData" / "WS" / "design"
            mutate(pdesign)
            pcfg = Cfg(branch_root=probe, workspace="WS", old_ch=78, display_offset=0)
            p = chapter_core.run_qa_validation(pcfg, quiet)
            w = openpyxl.load_workbook(p)
            rows = [(r[1].value, r[2].value, r[3].value)
                    for r in w["QA Validation"].iter_rows(min_row=2)]
            w.close()
            return rows

        def qa_fail_names(mutate):
            return [n for n, st, _ in qa_rows(mutate) if st == "FAIL"]

        def _break_hands(d):
            wb = openpyxl.load_workbook(d / "pve_level.xlsx")
            ws = wb["level"]
            for r in ws.iter_rows(min_row=2, max_col=7):
                if r[0].value == 79003:
                    r[6].value = 9          # outside the 4-7 range
            wb.save(d / "pve_level.xlsx")

        def _break_maxscore(d):
            wb = openpyxl.load_workbook(d / "pve_level.xlsx")
            ws = wb["level"]
            for r in ws.iter_rows(min_row=2, max_col=17):
                if r[0].value == 79004:
                    r[16].value = 1         # no longer 1.2 x 3★
            wb.save(d / "pve_level.xlsx")

        def _break_objective(d):
            wb = openpyxl.load_workbook(d / "pve_layout.xlsx")
            ws = wb["layout"]
            for r in ws.iter_rows(min_row=2, max_col=3):
                if r[0].value == 79005:
                    r[2].value = '{"card":3,}'   # collect-card objective unmet
            wb.save(d / "pve_layout.xlsx")

        def _break_switch(d):
            wb = openpyxl.load_workbook(d / "notification.xlsx")
            ws = wb["server_notification"]
            for r in ws.iter_rows(min_row=2, max_col=4):
                if r[1].value and "79章节" in str(r[1].value):
                    r[3].value = False      # new chapter left switched off
            wb.save(d / "notification.xlsx")

        for label, mutate, expect in [
            ("hands outside 4-7", _break_hands, "col G hands"),
            ("max score != 1.2 x 3★", _break_maxscore, "col Q max score"),
            ("notification left switched off", _break_switch, "switch (col D) is True"),
        ]:
            names = qa_fail_names(mutate)
            check(any(expect in n for n in names),
                  f"QA catches: {label}", f"FAILs were {names}")

        # The collect-card objective check is deliberately a WARNING, not a
        # gate: chapters that already shipped carry the same condition in bulk
        # (62 for internal chapter 93, 47 for 94), so failing on it blocked
        # sign-off for every chapter ever built.
        rows = qa_rows(_break_objective)
        obj = [(st, det) for n, st, det in rows if n and "收集卡" in n]
        check(len(obj) == 1, "the collect-card check appears exactly once",
              f"found {len(obj)}")
        if obj:
            st, det = obj[0]
            check(st == "WARN",
                  "QA warns rather than fails on an unmet collect-card objective",
                  f"status was {st}")
            check("already shipped" in (det or "") or "not present to compare" in (det or ""),
                  "the warning carries the previous chapter's count for context",
                  str(det)[:120])
        check(not any(st == "FAIL" and n and "收集卡" in n for n, st, _ in rows),
              "the collect-card condition never blocks sign-off")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

        # --- every workbook we save must be loaded formula-safe -------------
        # openpyxl writes a formula back without its cached value, and
        # coder_convert reads cached values - so a plain load -> edit -> save
        # blanks every formula cell. That is what wiped 96 aliases from
        # chest_conditions and made conversion fail with a check-ref error
        # pointing at 章节2星165, a row nothing to do with the chapter.
        # xlsx_safe.load_workbook_frozen exists for this; this guards against
        # a raw load_workbook creeping back into a save path.
        import re as _re
        src = Path(chapter_core.__file__).read_text(encoding="utf-8")
        src_lines = src.split("\n")
        offenders = []
        for i, line in enumerate(src_lines):
            m = _re.search(r"(\w+)\s*=\s*openpyxl\.load_workbook\(", line)
            if not m:
                continue
            var = m.group(1)
            window = src_lines[i:i + 160]
            if any(f"{var}.save(" in w for w in window):
                offenders.append(f"line {i+1}: {line.strip()[:70]}")
        check(not offenders,
              "no workbook is loaded with raw openpyxl and then saved",
              "; ".join(offenders[:3]))
        check("from xlsx_safe import load_workbook_frozen" in src,
              "chapter_core imports the formula-safe loader")

    print("")
    if FAILS:
        print(f"{len(FAILS)} CHECK(S) FAILED")
        sys.exit(1)
    print("CHAPTER CORE OK")


if __name__ == "__main__":
    main()
