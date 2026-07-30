import re

content = """
-- 活动 Item（只保留最临近的 2 次活动的）
--local ESim2507LocalService = require "Game/Module/LocalService/Event/ESim2507LocalService"
--ESim2507Bucket = ESim2507LocalService.New(LocalServiceMgr, "ESim2507Bucket"),
"""

tgt_evt = "Sim2507"
tgt_cap = tgt_evt.capitalize()
require_str_base = f'local E{tgt_cap}LocalService = require "Game/Module/LocalService/Event/E{tgt_cap}LocalService"'
bucket_str_base = f'E{tgt_cap}Bucket = E{tgt_cap}LocalService.New(LocalServiceMgr, "E{tgt_cap}Bucket"),'

commented_req_pattern = r'^\s*--\s*' + re.escape(require_str_base)
commented_bucket_pattern = r'^\s*--\s*(?:.*?)' + re.escape(bucket_str_base)

content = re.sub(commented_req_pattern, require_str_base, content, flags=re.MULTILINE)
content = re.sub(commented_bucket_pattern, r'        ' + bucket_str_base, content, flags=re.MULTILINE)

print(content)
