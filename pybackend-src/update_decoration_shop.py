import sys
import os
import re
import shutil
import stat
import json
import openpyxl
import copy as _copy


def copy_cell_format(src_cell, dst_cell):
    """Copy number format + style from src_cell to dst_cell (call after writing
    dst_cell.value). Plain value-only writes leave the destination at
    openpyxl's default "General" format, diverging from source cells
    formatted as e.g. Text."""
    dst_cell.number_format = src_cell.number_format
    if src_cell.has_style:
        dst_cell.font = _copy.copy(src_cell.font)
        dst_cell.fill = _copy.copy(src_cell.fill)
        dst_cell.border = _copy.copy(src_cell.border)
        dst_cell.alignment = _copy.copy(src_cell.alignment)
        dst_cell.protection = _copy.copy(src_cell.protection)


def shift_rows_down(ws, insert_row, num_new_rows, max_col=None):
    """Manually shift rows [insert_row, ws.max_row] down by num_new_rows,
    carrying value + format, then clear the vacated rows. Used instead of
    ws.insert_rows() so we keep full control/verification of format copy,
    matching the pattern already verified this session in asset_ref.xlsx /
    store.xlsx."""
    max_col = max_col or ws.max_column
    for r in range(ws.max_row, insert_row - 1, -1):
        for c in range(1, max_col + 1):
            src_c = ws.cell(r, c)
            dst_c = ws.cell(r + num_new_rows, c)
            dst_c.value = src_c.value
            copy_cell_format(src_c, dst_c)
    for r in range(insert_row, insert_row + num_new_rows):
        for c in range(1, max_col + 1):
            ws.cell(r, c).value = None


def find_divider_row(ws, marker, col=1, contains=False):
    """Find the row whose column `col` value matches `marker` (exact match by
    default, or substring match if contains=True). Returns None if not found."""
    for r in range(1, ws.max_row + 1):
        v = str(ws.cell(r, col).value or "")
        if (contains and marker in v) or (not contains and v.strip() == marker):
            return r
    return None


def last_row_with_id_prefix(ws, low, high, id_col=1):
    """Return the last row whose numeric id (column id_col, // stripped)
    falls in [low, high)."""
    last = None
    for r in range(2, ws.max_row + 1):
        raw = str(ws.cell(r, id_col).value or "").replace("//", "").strip()
        if raw.isdigit():
            n = int(raw)
            if low <= n < high:
                last = r
    return last


def get_real_last_row(ws, id_col=1):
    """ws.max_row can be inflated by phantom formatting on empty trailing rows
    (confirmed in item_cardBg: max_row=62 but row 62's id column is blank).
    Scan backward for the true last row with data in id_col."""
    for r in range(ws.max_row, 0, -1):
        v = ws.cell(r, id_col).value
        if v is not None and str(v).strip() != "":
            return r
    return 1


def last_row_matching_prefix(ws, col, prefix_regex):
    """Return (row, matched_number) for the last row whose column `col`
    matches prefix_regex (must have one capturing group, the trailing
    number)."""
    last_row = None
    last_num = None
    pattern = re.compile(prefix_regex)
    for r in range(2, ws.max_row + 1):
        v = str(ws.cell(r, col).value or "")
        m = pattern.match(v)
        if m:
            last_row = r
            last_num = int(m.group(1))
    return last_row, last_num


def main():
    try:
        design_path = sys.argv[1]
        payload_path = sys.argv[2]
    except IndexError:
        print("Error: Missing args")
        sys.exit(1)

    with open(payload_path, "r", encoding="utf-8-sig") as f:
        payload = json.load(f)

    event_id = str(payload["event_id"])
    event_name = payload["event_name"]  # e.g. "AliceInWonderland" — Game Mode 1's event
    game_mode_1 = payload["game_mode_1"]  # e.g. "Merge"
    game_mode_2 = payload["game_mode_2"]  # e.g. "Blocks"
    event_id_2 = str(payload["event_id_2"])  # Game Mode 2's own event, may differ from event_id
    event_name_2 = payload["event_name_2"]
    price_frame = payload.get("price_frame", "4")
    price_cardback = payload.get("price_cardback", "10")

    # Event ID is 8-digit yyyymmNN (e.g. 20260501) — alias uses YYMM, i.e. chars 2-6, not the leading century.
    id4 = event_id[2:6]
    id4_2 = event_id_2[2:6]
    alias_base = f"{id4}{event_name}"
    frame1_alias = f"{alias_base}{game_mode_1}"
    frame2_alias = f"{id4_2}{event_name_2}{game_mode_2}"
    # Card Back is always keyed off Game Mode 1's event.
    cardback_alias = alias_base

    icon_path = os.path.join(design_path, "icon.xlsx")
    item_path = os.path.join(design_path, "item.xlsx")
    asset_ref_path = os.path.join(design_path, "asset_ref.xlsx")
    card_path = os.path.join(design_path, "card.xlsx")
    skin_shop_path = os.path.join(design_path, "skin_shop.xlsx")

    for p in [icon_path, item_path, asset_ref_path, card_path, skin_shop_path]:
        if not os.path.exists(p):
            print(f"Error: required file not found: {p}")
            sys.exit(1)
        os.chmod(p, stat.S_IWRITE | stat.S_IREAD)

    try:
        # === 1. icon.xlsx -> icon_skin: 2 new frames + 1 new card skin (mid+big) ===
        wb_icon = openpyxl.load_workbook(icon_path)
        ws_icon = wb_icon["icon_skin"]

        # Frames: FRAME_SHOP_NN, inserted right before "// 卡牌皮肤" divider.
        last_frame_row, last_frame_num = last_row_matching_prefix(ws_icon, 1, r"^FRAME_SHOP_(\d+)$")
        card_skin_divider = find_divider_row(ws_icon, "// 卡牌皮肤")
        if card_skin_divider is None:
            raise ValueError("icon_skin: '// 卡牌皮肤' divider not found")
        frame_template = [ws_icon.cell(last_frame_row, c) for c in range(1, ws_icon.max_column + 1)]

        new_frame_nums = [last_frame_num + 1, last_frame_num + 2]
        shift_rows_down(ws_icon, card_skin_divider, 2)
        for i, (num, alias) in enumerate(zip(new_frame_nums, [frame1_alias, frame2_alias])):
            r = card_skin_divider + i
            values = [f"FRAME_SHOP_{num:02d}", alias, f"img_frame_skinshop{num}", "头像框100"]
            for c_idx, val in enumerate(values):
                dst_c = ws_icon.cell(r, c_idx + 1)
                dst_c.value = val
                copy_cell_format(frame_template[c_idx], dst_c)
        card_skin_divider += 2  # shifted down by the 2 rows we just inserted

        # Card skin (mid+big): CARD_SKIN_MID/BIG_NNN, inserted right before "//头像" divider.
        last_card_row, last_card_num = last_row_matching_prefix(ws_icon, 1, r"^CARD_SKIN_MID_(\d+)$")
        card_template = [ws_icon.cell(last_card_row, c) for c in range(1, ws_icon.max_column + 1)]
        card_big_template = [ws_icon.cell(last_card_row + 1, c) for c in range(1, ws_icon.max_column + 1)]
        avatar_divider = find_divider_row(ws_icon, "//头像")
        if avatar_divider is None:
            raise ValueError("icon_skin: '//头像' divider not found")
        new_card_num = last_card_num + 1
        shift_rows_down(ws_icon, avatar_divider, 2)
        mid_values = [f"CARD_SKIN_MID_{new_card_num:03d}", f"{cardback_alias}皮肤中图", f"img_cards_{new_card_num:03d}", "CARD_SKIN_60"]
        big_values = [f"CARD_SKIN_BIG_{new_card_num:03d}", f"{cardback_alias}皮肤大图", f"img_cardsshow_{new_card_num:03d}", "CARD_SKIN_B_60"]
        for c_idx, val in enumerate(mid_values):
            dst_c = ws_icon.cell(avatar_divider, c_idx + 1)
            dst_c.value = val
            copy_cell_format(card_template[c_idx], dst_c)
        for c_idx, val in enumerate(big_values):
            dst_c = ws_icon.cell(avatar_divider + 1, c_idx + 1)
            dst_c.value = val
            copy_cell_format(card_big_template[c_idx], dst_c)
        ws_icon.cell(avatar_divider + 1, 4).value = "CARD_SKIN_B_60"

        wb_icon.save(icon_path)
        print(f"icon.xlsx: added FRAME_SHOP_{new_frame_nums[0]}/{new_frame_nums[1]}, CARD_SKIN_MID/BIG_{new_card_num:03d}")

        # === 2. item.xlsx -> item: 2 new 5xxx (frames), 1 new 6xxx (card back) ===
        wb_item = openpyxl.load_workbook(item_path)
        ws_item = wb_item["item"]

        last_5xxx_row = last_row_with_id_prefix(ws_item, 5000, 6000)
        last_5xxx_id = int(str(ws_item.cell(last_5xxx_row, 1).value).replace("//", ""))
        item_5xxx_template = [ws_item.cell(last_5xxx_row, c) for c in range(1, ws_item.max_column + 1)]
        insert_5xxx_at = last_5xxx_row + 1
        shift_rows_down(ws_item, insert_5xxx_at, 2)
        new_5xxx_ids = [last_5xxx_id + 1, last_5xxx_id + 2]
        for i, (item_id, alias, frame_num) in enumerate(
            zip(new_5xxx_ids, [frame1_alias, frame2_alias], new_frame_nums)
        ):
            r = insert_5xxx_at + i
            row_vals = {1: item_id, 2: alias, 3: "头像框", 7: f"FRAME_SHOP_{frame_num:02d}"}
            for c in range(1, ws_item.max_column + 1):
                dst_c = ws_item.cell(r, c)
                dst_c.value = row_vals.get(c)
                copy_cell_format(item_5xxx_template[c - 1], dst_c)

        # 6xxx block shifted down by 2 already; recompute its true tail + the
        # "// 7000-7199" divider position dynamically (never hardcode rows).
        last_6xxx_row = last_row_with_id_prefix(ws_item, 6000, 7000)
        last_6xxx_id = int(str(ws_item.cell(last_6xxx_row, 1).value).replace("//", ""))
        item_6xxx_template = [ws_item.cell(last_6xxx_row, c) for c in range(1, ws_item.max_column + 1)]
        insert_6xxx_at = last_6xxx_row + 1
        divider_7xxx = find_divider_row(ws_item, "7000-7199", col=2, contains=True)
        if divider_7xxx is None:
            # fall back: the divider text lives in col1 in some files
            divider_7xxx = find_divider_row(ws_item, "7000-7199", col=1, contains=True)
        if divider_7xxx is not None and divider_7xxx != insert_6xxx_at:
            raise ValueError(
                f"item.xlsx: expected the '7000-7199' divider immediately after the 6xxx "
                f"block (row {insert_6xxx_at}) but found it at row {divider_7xxx} instead — "
                f"sheet layout has changed, stopping rather than risking an overwrite."
            )
        shift_rows_down(ws_item, insert_6xxx_at, 1)
        new_6xxx_id = last_6xxx_id + 1
        row_vals = {
            1: new_6xxx_id, 2: cardback_alias, 3: "装饰卡背卡面",
            7: f"{cardback_alias}皮肤中图", 8: f"{cardback_alias}皮肤中图", 9: f"{cardback_alias}皮肤大图",
        }
        for c in range(1, ws_item.max_column + 1):
            dst_c = ws_item.cell(insert_6xxx_at, c)
            dst_c.value = row_vals.get(c)
            copy_cell_format(item_6xxx_template[c - 1], dst_c)

        print(f"item.xlsx [item]: added 5xxx {new_5xxx_ids}, 6xxx {new_6xxx_id}")

        # === 3. item.xlsx -> item_frame: append 2 rows (true tail, no adjacent section) ===
        ws_if = wb_item["item_frame"]
        if_last_row = get_real_last_row(ws_if)
        if_template = [ws_if.cell(if_last_row, c) for c in range(1, ws_if.max_column + 1)]
        insert_at = if_last_row + 1
        for i, (item_id, alias) in enumerate(zip(new_5xxx_ids, [frame1_alias, frame2_alias])):
            r = insert_at + i
            # col 7 (获取参数) stays blank for Deco_Shop rows (only Star_Pass etc. use it);
            # col 8 (不支持预览) is "=TRUE()", matching every existing Deco_Shop row.
            row_vals = {1: item_id, 2: alias, 3: "DECO_PANEL_TIP_DECO_SHOP", 6: "Deco_Shop", 8: "=TRUE()"}
            for c in range(1, ws_if.max_column + 1):
                dst_c = ws_if.cell(r, c)
                dst_c.value = row_vals.get(c)
                copy_cell_format(if_template[c - 1], dst_c)

        # === 4. item.xlsx -> item_cardBg: append 1 row (true tail) ===
        ws_icb = wb_item["item_cardBg"]
        icb_last_row = get_real_last_row(ws_icb)
        icb_template = [ws_icb.cell(icb_last_row, c) for c in range(1, ws_icb.max_column + 1)]
        r = icb_last_row + 1
        # col 6 (获取方式)/col 8 (不支持预览) match every existing Deco_Shop row;
        # col 7 (获取参数) stays blank, same as item_frame.
        row_vals = {
            1: new_6xxx_id, 2: cardback_alias, 3: cardback_alias, 4: "DECO_PANEL_TIP_DECO_SHOP",
            6: "Deco_Shop", 8: "=TRUE()",
        }
        for c in range(1, ws_icb.max_column + 1):
            dst_c = ws_icb.cell(r, c)
            dst_c.value = row_vals.get(c)
            copy_cell_format(icb_template[c - 1], dst_c)

        wb_item.save(item_path)
        print("item.xlsx: item_frame + item_cardBg rows appended")

        # === 5. asset_ref.xlsx -> key_asset: 1 new CARD_SHOP_NN before "// StarPass" ===
        wb_asset = openpyxl.load_workbook(asset_ref_path)
        ws_key = wb_asset["key_asset"]
        last_cardshop_row, last_cardshop_num = last_row_matching_prefix(ws_key, 1, r"^CARD_SHOP_(\d+)$")
        key_template = [ws_key.cell(last_cardshop_row, c) for c in range(1, ws_key.max_column + 1)]
        starpass_divider = find_divider_row(ws_key, "// StarPass")
        if starpass_divider is None:
            raise ValueError("key_asset: '// StarPass' divider not found")
        new_cardshop_num = last_cardshop_num + 1
        shift_rows_down(ws_key, starpass_divider, 1)
        cardshop_name = f"CARD_SHOP_{new_cardshop_num}"
        row_vals = {1: cardshop_name, 2: cardshop_name, 3: f"Assets/ResourcesAb/Icon/Card_SkinShop_{new_cardshop_num}"}
        for c in range(1, ws_key.max_column + 1):
            dst_c = ws_key.cell(starpass_divider, c)
            dst_c.value = row_vals.get(c)
            copy_cell_format(key_template[c - 1], dst_c)
        wb_asset.save(asset_ref_path)
        print(f"asset_ref.xlsx: added {cardshop_name}")

        # === 6. card.xlsx -> card_skin_atlas: append 1 row (true tail) ===
        wb_card = openpyxl.load_workbook(card_path)
        ws_atlas = wb_card["card_skin_atlas"]
        last_base = 0
        last_base_row = None
        for r in range(2, ws_atlas.max_row + 1):
            v = ws_atlas.cell(r, 1).value
            if isinstance(v, (int, float)) and v > last_base:
                last_base = int(v)
                last_base_row = r
        new_base = last_base + 10000
        atlas_template = [ws_atlas.cell(last_base_row, c) for c in range(1, ws_atlas.max_column + 1)]
        r = get_real_last_row(ws_atlas) + 1
        row_vals = {1: new_base, 2: cardback_alias, 3: cardshop_name, 4: cardshop_name}
        for c in range(1, ws_atlas.max_column + 1):
            dst_c = ws_atlas.cell(r, c)
            dst_c.value = row_vals.get(c)
            copy_cell_format(atlas_template[c - 1], dst_c)

        # === 7. card.xlsx -> card_style: clone a skin's 18-row block (offsets
        # 0,1-13,28-31 — no legacy 18th "随机SkipBo卡背" row, that mechanic was
        # discontinued after base 440000, confirmed against every skin since). ===
        ws_style = wb_card["card_style"]
        template_offsets = [0] + list(range(1, 14)) + [28, 29, 30, 31]

        def rows_for_base(base):
            rows = []
            for r in range(2, ws_style.max_row + 1):
                idv = ws_style.cell(r, 1).value
                if isinstance(idv, (int, float)) and (idv // 10000) * 10000 == base:
                    rows.append(r)
            return rows

        # The immediately-previous skin can be a one-off special-art event with its
        # own dedicated per-skin card-back images (e.g. alice2608's "alice2608皮肤局内
        # 卡背N", registered only for that skin) instead of the shared generic asset
        # ("皮肤配置卡面通用..."). Cloning that verbatim and renaming the token would
        # reference an image nobody registered for the new skin. Every skin back to
        # base 10000 uses the generic asset except that kind of one-off — so walk
        # backward to the nearest base whose col7 is generic before using it as the
        # content template. ID numbering (new_base) still continues from last_base.
        template_base = last_base
        template_rows = rows_for_base(template_base)
        while template_rows and not str(ws_style.cell(template_rows[0], 7).value or "").startswith("皮肤配置"):
            template_base -= 10000
            template_rows = rows_for_base(template_base)

        if len(template_rows) != len(template_offsets):
            raise ValueError(
                f"card_style: expected {len(template_offsets)} rows for template base "
                f"{template_base}, found {len(template_rows)} — refusing to guess, check the sheet."
            )
        old_alias_prefix = str(ws_style.cell(template_rows[0], 2).value)
        # Strip the trailing "卡面1..." descriptor to find the bare skin-name token to swap.
        m = re.match(r"^(.*?)(卡面1|卡背1).*$", old_alias_prefix)
        old_token = m.group(1) if m else old_alias_prefix

        insert_at = get_real_last_row(ws_style) + 1
        for i, (offset, src_r) in enumerate(zip(template_offsets, template_rows)):
            dst_r = insert_at + i
            for c in range(1, ws_style.max_column + 1):
                src_c = ws_style.cell(src_r, c)
                val = src_c.value
                if c == 1:
                    val = new_base + offset
                elif isinstance(val, str) and old_token:
                    val = val.replace(old_token, cardback_alias)
                dst_c = ws_style.cell(dst_r, c)
                dst_c.value = val
                copy_cell_format(src_c, dst_c)

        wb_card.save(card_path)
        print(f"card.xlsx: card_skin_atlas base {new_base}, card_style {len(template_offsets)} rows cloned")

        # === 8. skin_shop.xlsx -> Skin Item: 1 new 10xxx (card back), 2 new 20xxx (frames) ===
        wb_shop = openpyxl.load_workbook(skin_shop_path)
        ws_shopitem = wb_shop["皮肤Item"]

        # 10xxx block: find last numeric ID starting with "100" (card-back series)
        last_10xxx_row = None
        last_10xxx_id = None
        for r in range(2, ws_shopitem.max_row + 1):
            raw = str(ws_shopitem.cell(r, 1).value or "").strip()
            if raw.isdigit() and raw.startswith("100"):
                last_10xxx_row = r
                last_10xxx_id = int(raw)
        shopitem_10xxx_template = [ws_shopitem.cell(last_10xxx_row, c) for c in range(1, ws_shopitem.max_column + 1)]
        frame_series_divider = find_divider_row(ws_shopitem, "// 头像框系列")
        if frame_series_divider is None:
            raise ValueError("皮肤Item: '// 头像框系列' divider not found")
        if frame_series_divider != last_10xxx_row + 1:
            raise ValueError(
                f"皮肤Item: expected '// 头像框系列' right after the last 10xxx row "
                f"({last_10xxx_row}) but found it at {frame_series_divider} — layout changed, stopping."
            )
        new_10xxx_id = last_10xxx_id + 1
        shift_rows_down(ws_shopitem, frame_series_divider, 1)
        # col 1 (皮肤商品ID) must be a string, matching every existing row —
        # a raw int here breaks the client's ID lookup against the 皮肤商店 shelf list.
        row_vals = {1: str(new_10xxx_id), 3: "卡牌", 4: str(new_6xxx_id), 6: price_cardback}
        for c in range(1, ws_shopitem.max_column + 1):
            dst_c = ws_shopitem.cell(frame_series_divider, c)
            dst_c.value = row_vals.get(c)
            copy_cell_format(shopitem_10xxx_template[c - 1], dst_c)

        # 20xxx frames: true tail, safe to append.
        last_20xxx_row = None
        last_20xxx_id = None
        for r in range(2, ws_shopitem.max_row + 1):
            raw = str(ws_shopitem.cell(r, 1).value or "").strip()
            if raw.isdigit() and raw.startswith("200"):
                last_20xxx_row = r
                last_20xxx_id = int(raw)
        shopitem_20xxx_template = [ws_shopitem.cell(last_20xxx_row, c) for c in range(1, ws_shopitem.max_column + 1)]
        insert_at = get_real_last_row(ws_shopitem) + 1
        new_20xxx_ids = [last_20xxx_id + 1, last_20xxx_id + 2]
        for i, (sku_id, item_id) in enumerate(zip(new_20xxx_ids, new_5xxx_ids)):
            r = insert_at + i
            row_vals = {1: str(sku_id), 3: "头像框", 4: str(item_id), 6: price_frame}
            for c in range(1, ws_shopitem.max_column + 1):
                dst_c = ws_shopitem.cell(r, c)
                dst_c.value = row_vals.get(c)
                copy_cell_format(shopitem_20xxx_template[c - 1], dst_c)

        # === 9. skin_shop.xlsx -> Skin Store (皮肤商店): rotate the fixed-size shelf ===
        ws_store = wb_shop["皮肤商店"]
        for r in range(2, ws_store.max_row + 1):
            alias = str(ws_store.cell(r, 2).value or "")
            cell = ws_store.cell(r, 6)
            ids = [x.strip() for x in str(cell.value or "").split(",") if x.strip()]
            if alias == "卡背":
                ids = ids[1:] + [str(new_10xxx_id)]  # drop oldest, append newest
            elif alias == "头像框":
                ids = ids[2:] + [str(x) for x in new_20xxx_ids]  # drop 2 oldest, append 2 newest
            cell.value = ",".join(ids)

        wb_shop.save(skin_shop_path)
        print(f"skin_shop.xlsx: added 10xxx={new_10xxx_id}, 20xxx={new_20xxx_ids}, rotated shelf")

        # === 10. Duplicate image assets as placeholders (artist replaces later) ===
        def dup_asset(folder, old_name, new_name):
            src = os.path.join(folder, old_name)
            if not os.path.exists(src):
                print(f"WARNING: source asset not found, skipping copy: {src}")
                return
            dst = os.path.join(folder, new_name)
            if os.path.exists(dst):
                # An artist may have already pre-positioned real art at this filename
                # ahead of the data entry catching up — never clobber it with a placeholder.
                print(f"WARNING: destination asset already exists, skipping copy (not overwriting real art): {dst}")
                return
            shutil.copy2(src, dst)
            meta_src = src + ".meta"
            if os.path.exists(meta_src):
                shutil.copy2(meta_src, dst + ".meta")

        # design_path = .../<branch>/design/DesignData/<workspace>/design
        # -> branch root is 4 levels up.
        branch_root = design_path
        for _ in range(4):
            branch_root = os.path.dirname(branch_root)
        client_root = os.path.join(branch_root, "develop", "client", "Skipbo", "Assets", "ResourcesAb", "Icon")
        client_root = os.path.normpath(client_root)
        dup_asset(os.path.join(client_root, "HeadFrame100"), f"img_frame_skinshop{last_frame_num}.png", f"img_frame_skinshop{new_frame_nums[0]}.png")
        dup_asset(os.path.join(client_root, "HeadFrame100"), f"img_frame_skinshop{last_frame_num}.png", f"img_frame_skinshop{new_frame_nums[1]}.png")
        dup_asset(os.path.join(client_root, "CardShow60"), f"img_cards_{last_card_num:03d}.png", f"img_cards_{new_card_num:03d}.png")
        dup_asset(os.path.join(client_root, "CardShowB60"), f"img_cardsshow_{last_card_num:03d}.png", f"img_cardsshow_{new_card_num:03d}.png")

        print("Successfully updated Decoration Shop configuration.")

    except Exception as e:
        print(f"Fatal Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
