import json
from pathlib import Path

from app import solve_schedule

BASE = Path(__file__).resolve().parent
INPUT_DIR = BASE / 'tests' / 'inputs'
MANIFEST_DIR = BASE / 'tests' / 'manifests'
GENERATED_DIR = BASE / 'tests' / 'generated_outputs'


def assert_invariants(input_data: dict, result: dict) -> None:
    assert result['success'] is True
    config = input_data['config']
    max_cap = config.get('max_shifts_per_member', 1)
    members = {m['name']: m for m in input_data['members']}
    seen_shift = set()
    counts = {name: 0 for name in members}
    for item in result['assignments']:
        shift = (item['day'], item['slot'])
        person = item['person']
        assert shift not in seen_shift, f'班次重复分配: {shift}'
        seen_shift.add(shift)
        assert config['open_shifts'][item['day']][item['slot']] == 1, f'分配到了不开班时段: {item}'
        assert person in members, f'未知成员: {person}'
        member = members[person]
        assert item['day'] not in member.get('ban_days', []), f'分配到了禁排日期: {item}'
        assert member['schedule'][item['day']][item['slot']] == 0, f'分配到了有课时段: {item}'
        counts[person] += 1
        assert counts[person] <= max_cap, f'成员超过用户设置上限: {person}'
    assert result['summary']['assigned_shift_count'] == len(result['assignments'])
    assert result['summary']['unfilled_shift_count'] == len(result['unfilled_shifts'])
    assert result['summary']['unassigned_member_count'] == len(result['unassigned_members'])
    assert result['summary']['max_shifts_used_in_result'] == (max(counts.values()) if counts else 0)


def run_one(manifest_path: Path) -> None:
    manifest = json.loads(manifest_path.read_text(encoding='utf-8'))
    input_path = INPUT_DIR / manifest['input']
    data = json.loads(input_path.read_text(encoding='utf-8'))
    result = solve_schedule(data)
    assert_invariants(data, result)
    expected = manifest['expected']
    for key, value in expected.items():
        if key.startswith('summary.'):
            field = key.split('.', 1)[1]
            assert result['summary'][field] == value, f'{field}: expected {value}, got {result["summary"][field]}'
        elif key.startswith('capacity_advice.'):
            field = key.split('.', 1)[1]
            assert result['capacity_advice'][field] == value, f'{field}: expected {value}, got {result["capacity_advice"][field]}'
        else:
            assert result[key] == value, f'{key}: expected {value}, got {result[key]}'
    GENERATED_DIR.mkdir(parents=True, exist_ok=True)
    out_path = GENERATED_DIR / (manifest_path.stem.replace('.manifest', '') + '.verified.json')
    out_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'[PASS] {manifest_path.name}')


if __name__ == '__main__':
    manifests = sorted(MANIFEST_DIR.glob('*.json'))
    if not manifests:
        raise SystemExit('没有找到测试 manifest')
    for manifest in manifests:
        run_one(manifest)
    print(f'共通过 {len(manifests)} 组测试。')
