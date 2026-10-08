"""Self-contained JSON plans: ordered operations plus immutable rule snapshots."""
import copy
import json
import os

from .pipeline import PipelineError, normalize_recipe


def _read(path):
    with open(path, encoding='utf-8-sig') as handle:
        return json.load(handle)


def _replacement_path(name, base):
    if os.path.isabs(name) and os.path.isfile(name):
        return name
    for candidate in (os.path.join(base, 'replace-rules', name+'.json'),
                      os.path.join(base, name+'.json'), os.path.join(base, name)):
        if os.path.isfile(candidate): return candidate
    raise PipelineError('方案引用的替换规则不存在：%s。请先保存这套规则再导出方案。' % name)


def validate(data):
    if not isinstance(data, dict) or not isinstance(data.get('steps'), list):
        raise PipelineError('方案文件需要包含 steps 步骤清单')
    if not data['steps']: raise PipelineError('方案没有处理步骤')
    normalize_recipe(data['steps'])
    for item in data['steps']:
        if not isinstance(item, dict): continue
        params = item.get('params') or {}
        if not isinstance(params, dict): raise PipelineError('步骤参数必须是对象')
        inline = params.get('rules_data')
        if inline is not None and not isinstance(inline, dict):
            raise PipelineError('方案内的规则内容不是有效对象')
        if item.get('op') == 'sup' and inline is not None:
            from .rules import RuleSet
            issues = RuleSet(inline).validate()
            if issues: raise PipelineError('上下标字典有误：%s' % '；'.join(issues))
        if item.get('op') == 'tablestyle' and params.get('styles_data') is not None:
            from .tablestyle import StyleSet
            styles = StyleSet(params['styles_data'])
            requested = [params.get('style')]+[row.get('style') for row in params.get('mapping') or [] if not row.get('none')]
            missing = [name for name in requested if name and name not in styles.styles]
            if missing: raise PipelineError('方案缺少表格款式：%s' % '、'.join(missing))
    return copy.deepcopy(data)


def export(data, rules_base):
    result = validate(data)
    result['format'] = 'wordfactory-plan'
    result['schema'] = 2
    steps = []
    for original in result['steps']:
        item = copy.deepcopy(original if isinstance(original,dict) else {'op':original})
        params = item.setdefault('params', {})
        op = item.get('op')
        if op == 'replace' and params.get('rules') and 'rules_data' not in params:
            params['rules_data'] = _read(_replacement_path(params['rules'], rules_base))
        elif op == 'sup' and 'rules_data' not in params:
            path = params.get('rules') or os.path.join(rules_base,'subscripts.json')
            if not os.path.isabs(path): path = os.path.join(rules_base,path)
            params['rules_data'] = _read(path)
        elif op == 'tablestyle' and 'styles_data' not in params:
            params['styles_data'] = _read(os.path.join(rules_base,'tablestyle.json'))
        steps.append(item)
    result['steps'] = steps
    if not result.get('font_rules_data'):
        result['font_rules_data'] = _read(os.path.join(rules_base,'fonts.json'))
    return validate(result)


def import_plan(data):
    result = validate(data)
    if result.get('format') not in (None,'wordfactory-plan'):
        raise PipelineError('这不是 word工厂 的执行方案文件')
    if int(result.get('schema',1)) > 2:
        raise PipelineError('方案版本比本程序新，请更新 word工厂 后导入')
    return result
