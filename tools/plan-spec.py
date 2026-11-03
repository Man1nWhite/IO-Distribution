#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Исполнимая спецификация укладчика реек (`TryPlanBlocks` и компания).

Зеркало планировщика из `Module1.bas`: `TryPlanBlocks`, `SearchBlocks`
(включая отсечение симметрии одинаковых по ширине блоков и дедупликацию
реек с равным свободным местом), `TryPlaceGroupInCab`,
`TryPlacePartialSplit` (фазы 1 и 2; фаза 2 начинает со второго блока — её
первый вариант повторял бы первый вариант фазы 1), `CommitPlan` и
`РазместитьСвязкуНаСвободнойРейке` (этап 1.1, шкаф P410). Порядок обхода,
сортировка блоков по убыванию и выбор реек повторяют VBA; размещение
фиксируется только после успешной укладки всей группы (как `CommitPlan`),
неудачная попытка состояние реек не меняет.

Сквозной бюджет узлов: один счётчик `Budget` на попытку
`try_place_group_in_cab` — общий для целых блоков, мягкого дробления и
фоллбэка на отдельные элементы, предел `PLAN_NODE_LIMIT` (в VBA счётчик
живёт в `TryPlaceGroupInCab` и пробрасывается ByRef).

Это НЕ тест VBA (Excel/VBE здесь нет), а канон поведения планировщика:
канонические сценарии, проверка инвариантов (нет переполнения рейки, все
элементы размещены, связка не рвётся между шкафами) и рандомизированные
прогоны с фиксированным seed со сверкой с брутфорсом на малых задачах.
Прогонов два: круглые ширины (равные остатки реек — проверка дедупликации
в SearchBlocks) и настоящие ширины модулей (13, 23, 157, …). Каждый
прогон проверяет и то, что он не выродился: лимит перебора не достигнут
(иначе сравнивать с брутфорсом нечего) и есть как укладки, так и отказы.

МЕНЯЕШЬ АЛГОРИТМ В VBA — МЕНЯЙ ЭТОТ ФАЙЛ. Автоматической сверки нет:
зеркало переводится из VBA вручную, расхождение смотрит человек.

Код возврата: 0 — все проверки сошлись, 1 — есть расхождение.
"""

import random
import sys
import time

EPS = 0.0001            # PLAN_EPS
NODE_LIMIT = 300000     # PLAN_NODE_LIMIT
SEARCH_BLOCK_LIMIT = 120  # PLAN_SEARCH_BLOCK_LIMIT


class Budget:
    """Сквозной счётчик узлов перебора на одну попытку группы.

    search_calls / plan_calls / variants — инструменты проверки, а не
    состояние VBA: по ним видно, сколько поисков потрачено и не пробовался
    ли один и тот же набор блоков дважды.
    """

    def __init__(self, limit=NODE_LIMIT):
        self.limit = limit
        self.nodes = 0
        self.search_calls = 0
        self.plan_calls = 0
        self.variants = []


class Rails:
    """Состояние реек: (cab, kind, cap, used, список элементов)."""

    def __init__(self, spec):
        self.cab = [s[0] for s in spec]
        self.kind = [s[1] for s in spec]
        self.cap = [float(s[2]) for s in spec]
        self.used = [float(s[3]) if len(s) > 3 else 0.0 for s in spec]
        self.base = list(self.used)  # занято до нас: в used, но не в elems
        self.elems = [[] for _ in spec]

    def __len__(self):
        return len(self.cap)

    def snapshot(self):
        return (list(self.used), [list(e) for e in self.elems])

    def cab_rails(self, cab_id, kind):
        return [i for i in range(len(self))
                if self.cab[i] == cab_id and self.kind[i] == kind]


# ---------------------------------------------------------------------------
# Зеркало TryPlanBlocks / SearchBlocks. plan: блок -> индекс рейки в rails.
# Ничего не размещает; при неудаче возвращает (None, hit_limit).
# ---------------------------------------------------------------------------

def search_blocks(level, n_blk, order, blk_len, rail_ids, free_cap,
                  plan, budget, hit):
    """Рекурсивный перебор; hit — список из одного bool (ByRef в VBA)."""
    if hit[0]:
        return False
    budget.nodes += 1
    if budget.nodes > budget.limit:
        hit[0] = True
        return False
    if level > n_blk:
        return True
    blk = order[level - 1]
    start = 0
    if level > 1:
        prev = order[level - 2]
        if abs(blk_len[blk] - blk_len[prev]) <= EPS:
            for i, rid in enumerate(rail_ids):
                if rid == plan[prev]:
                    start = i
                    break
    seen_free = []
    for i in range(start, len(rail_ids)):
        if free_cap[i] + EPS >= blk_len[blk]:
            if any(abs(free_cap[i] - f) <= EPS for f in seen_free):
                continue
            seen_free.append(free_cap[i])
            free_cap[i] -= blk_len[blk]
            plan[blk] = rail_ids[i]
            if search_blocks(level + 1, n_blk, order, blk_len, rail_ids,
                             free_cap, plan, budget, hit):
                return True
            plan[blk] = -1
            free_cap[i] += blk_len[blk]
            if hit[0]:
                return False
    return False


def try_plan_blocks(rails, lens, cab_id, kind, items, blocks, budget):
    """(plan|None, hit_limit). items/blocks — параллельные списки элементов."""
    budget.plan_calls += 1
    budget.variants.append(tuple(blocks))
    n_blk = max(blocks) + 1 if blocks else 0
    blk_len = [0.0] * n_blk
    total = 0.0
    for idx, blk in zip(items, blocks):
        blk_len[blk] += lens[idx]
        total += lens[idx]
    rail_ids = rails.cab_rails(cab_id, kind)
    if not rail_ids:
        return None, False
    free_cap = [max(0.0, rails.cap[i] - rails.used[i]) for i in rail_ids]
    if total > sum(free_cap) + EPS:
        return None, False

    # Приоритет — вся группа на одной рейке (наименьший остаток).
    best, best_waste = -1, 1e30
    for i in range(len(rail_ids)):
        waste = free_cap[i] - total
        if waste >= -EPS and waste < best_waste:
            best, best_waste = i, waste
    if best >= 0:
        return {b: rail_ids[best] for b in range(n_blk)}, False

    if any(bl > max(free_cap) + EPS for bl in blk_len):
        return None, False

    # Сортировка блоков по убыванию длины (устойчивая, как в VBA).
    order = sorted(range(n_blk), key=lambda b: -blk_len[b])

    # Best-fit decreasing.
    plan = {}
    free_greedy = list(free_cap)
    greedy_ok = True
    for blk in order:
        best, best_waste = -1, 1e30
        for i in range(len(rail_ids)):
            waste = free_greedy[i] - blk_len[blk]
            if waste >= -EPS and waste < best_waste:
                best, best_waste = i, waste
        if best < 0:
            greedy_ok = False
            break
        plan[blk] = rail_ids[best]
        free_greedy[best] -= blk_len[blk]
    if greedy_ok:
        return plan, False

    if n_blk > SEARCH_BLOCK_LIMIT:
        return None, True

    budget.search_calls += 1
    hit = [False]
    plan = {b: -1 for b in range(n_blk)}
    if search_blocks(1, n_blk, order, blk_len, rail_ids, list(free_cap),
                     plan, budget, hit):
        return plan, False
    return None, hit[0]


def commit_plan(rails, lens, placed_rail, items, blocks, plan):
    """Зеркало CommitPlan: размещает все элементы по плану блоков."""
    for idx, blk in zip(items, blocks):
        ri = plan[blk]
        rails.used[ri] += lens[idx]
        rails.elems[ri].append(idx)
        placed_rail[idx] = ri


def try_place_partial_split(rails, lens, placed_rail, cab_id, kind,
                            items, blocks, budget):
    """Мягкое дробление. (ok, hit_limit); размещает только при успехе."""
    n_blk = max(blocks) + 1
    if n_blk < 2:
        return False, False
    blk_cnt = [0] * n_blk
    blk_len = [0.0] * n_blk
    for idx, blk in zip(items, blocks):
        blk_cnt[blk] += 1
        blk_len[blk] += lens[idx]

    for phase in (1, 2):
        done = [False] * n_blk
        tried = [False] * n_blk
        for n_try in range(n_blk):
            best = -1
            for blk in range(n_blk):
                if blk_cnt[blk] > 1 and not tried[blk]:
                    if best < 0 or blk_len[blk] > blk_len[best] + EPS:
                        best = blk
            if best < 0:
                break
            tried[best] = True
            if phase == 1:
                done = [False] * n_blk
            done[best] = True

            # Фаза 2, шаг 1 — тот же вариант, что шаг 1 фазы 1 (раздроблен
            # только самый широкий блок). Повтор тратит сквозной бюджет
            # узлов, не меняя исход: фаза 2 начинается со второго блока.
            if phase == 2 and n_try == 0:
                continue

            blk_map = {}
            nxt = 0
            for blk in range(n_blk):
                if blk_cnt[blk] > 0 and not done[blk]:
                    blk_map[blk] = nxt
                    nxt += 1
            part = []
            part_n = nxt
            for blk in blocks:
                if done[blk]:
                    part_n += 1
                    part.append(part_n - 1)
                else:
                    part.append(blk_map[blk])
            if part_n >= len(items):
                break  # целых блоков не осталось — штатный фоллбэк
            plan, hit = try_plan_blocks(rails, lens, cab_id, kind,
                                        items, part, budget)
            if plan is not None:
                commit_plan(rails, lens, placed_rail, items, part, plan)
                return True, False
            if hit:
                return False, True
    return False, False


def try_place_group_in_cab(rails, lens, placed_rail, cab_id, kind,
                           items, blocks, budget=None):
    """Попытка группы в один шкаф. -> dict(ok, was_split, hit_limit, nodes)."""
    if budget is None:
        budget = Budget()
    plan, hit = try_plan_blocks(rails, lens, cab_id, kind,
                                items, blocks, budget)
    if plan is not None:
        commit_plan(rails, lens, placed_rail, items, blocks, plan)
        return {"ok": True, "was_split": False, "hit_limit": False,
                "nodes": budget.nodes}
    hit_limit = hit
    was_split = False
    if max(blocks) + 1 < len(items):
        ok, hit2 = try_place_partial_split(rails, lens, placed_rail,
                                           cab_id, kind, items, blocks,
                                           budget)
        if ok:
            return {"ok": True, "was_split": True,
                    "hit_limit": hit_limit, "nodes": budget.nodes}
        hit_limit = hit_limit or hit2
        singles = list(range(len(items)))
        plan, hit = try_plan_blocks(rails, lens, cab_id, kind,
                                    items, singles, budget)
        if plan is not None:
            commit_plan(rails, lens, placed_rail, items, singles, plan)
            return {"ok": True, "was_split": True,
                    "hit_limit": hit_limit, "nodes": budget.nodes}
        hit_limit = hit_limit or hit
    return {"ok": False, "was_split": was_split, "hit_limit": hit_limit,
            "nodes": budget.nodes}


def place_bundle_on_free_rail(rails, lens, placed_rail, cab_id, kind, items):
    """РазместитьСвязкуНаСвободнойРейке: связка целиком, рейка с максимумом
    свободного места; нет такой — False и состояние не меняется."""
    total = sum(lens[i] for i in items)
    best, best_free = -1, -1.0
    for i in rails.cab_rails(cab_id, kind):
        cur = max(0.0, rails.cap[i] - rails.used[i])
        if cur + EPS >= total and cur > best_free:
            best, best_free = i, cur
    if best < 0:
        return False
    for idx in items:
        rails.used[best] += lens[idx]
        rails.elems[best].append(idx)
        placed_rail[idx] = best
    return True


# ---------------------------------------------------------------------------
# Брутфорс (эталон на малых задачах) и инварианты.
# ---------------------------------------------------------------------------

def brute_force(lengths, caps):
    """Можно ли разложить неделимые куски lengths по рейкам caps.

    Сначала дешёвый отказ: не хватает суммарной ёмкости или самый длинный
    кусок длиннее самой длинной рейки. Без этого заведомо невозможная
    задача обходила всё дерево — при 16 элементах и четырёх рейках это
    самая дорогая ветка перебора.

    Допуск суммы — EPS на элемент, а не один EPS: go() прощает до EPS на
    каждое размещение, поэтому одиночный допуск отсекал бы укладки,
    которые перебор считает возможными. С множителем условие строго
    слабее перебора и ответ не меняет.
    """
    lengths = sorted(lengths, reverse=True)
    if sum(lengths) > sum(caps) + EPS * max(1, len(lengths)):
        return False
    if lengths and lengths[0] > max(caps, default=0) + EPS:
        return False
    free = list(caps)

    def go(k):
        if k == len(lengths):
            return True
        seen = set()
        for i in range(len(free)):
            if free[i] + EPS >= lengths[k] and free[i] not in seen:
                seen.add(free[i])
                free[i] -= lengths[k]
                if go(k + 1):
                    return True
                free[i] += lengths[k]
        return False

    return go(0)


def check_invariants(name, rails, lens, placed_rail, groups_cabs):
    """groups_cabs: список (элементы связки/группы, cab). -> список проблем."""
    problems = []
    for i in range(len(rails)):
        if rails.used[i] > rails.cap[i] + EPS * (len(rails.elems[i]) + 1):
            problems.append(f"{name}: рейка {i} переполнена "
                            f"({rails.used[i]} > {rails.cap[i]})")
        s = sum(lens[e] for e in rails.elems[i])
        if abs(s - (rails.used[i] - rails.base[i])) > \
                EPS * (len(rails.elems[i]) + 1):
            problems.append(f"{name}: рейка {i}: used != сумме ширин")
    seen = set()
    for items, cab_id in groups_cabs:
        for idx in items:
            if idx in seen:
                problems.append(f"{name}: элемент {idx} размещён дважды")
            seen.add(idx)
            if idx not in placed_rail:
                problems.append(f"{name}: элемент {idx} не размещён")
            elif rails.cab[placed_rail[idx]] != cab_id:
                problems.append(f"{name}: элемент {idx} в чужом шкафу "
                                f"(связка разорвана между шкафами)")
    return problems


FAIL = []


def check(name, cond, detail=""):
    print(("ok   " if cond else "FAIL ") + name + (f"  {detail}" if detail and not cond else ""))
    if not cond:
        FAIL.append(name)


def intact_bundles(blocks, items, placed_rail, skip=()):
    """Сколько блоков лежат целиком на одной рейке (кроме блоков из skip).

    skip — блоки, которые не считаются связками: в сценарии DIO блок 0 —
    сам модуль DIO, он всегда целый и в счёт «7 из 8 связок» не входит.
    """
    by_block = {}
    for idx, blk in zip(items, blocks):
        if blk in skip:
            continue
        by_block.setdefault(blk, set()).add(placed_rail[idx])
    return sum(1 for rs in by_block.values() if len(rs) == 1)


# ---------------------------------------------------------------------------
# Канонические сценарии.
# ---------------------------------------------------------------------------

def scenario_p410_bundles():
    """4 KDI + 4 KDO в шкафу P410: связки целы, раскладка F/F/B/B = 1/1/3/3."""
    # Рейки F,F,B,B; фронтальные урезаны зоной P410 на 224 мм (D18, 2000 мм).
    rails = Rails([(1, "M", 2000 - 224), (1, "M", 2000 - 224),
                   (1, "M", 2000), (1, "M", 2000)])
    lens, placed, groups = [], {}, []
    for _ in range(8):  # связка: реле 110 + 4 барьера по 13
        items = list(range(len(lens), len(lens) + 5))
        lens += [110.0, 13.0, 13.0, 13.0, 13.0]
        check("P410: связка целиком на свободную рейку",
              place_bundle_on_free_rail(rails, lens, placed, 1, "M", items))
        groups.append((items, 1))
    counts = [len(rails.elems[i]) // 5 for i in range(len(rails))]
    # Фронтальные рейки урезаны на 224 мм, поэтому при равных связках задние
    # (полные 2000 мм) заполняются плотнее фронтальных: [F, F, B, B].
    check("P410: все связки целы, раскладка F/F/B/B = 1/1/3/3",
          counts == [1, 1, 3, 3], str(counts))
    probs = check_invariants("P410", rails, lens, placed, groups)
    check("P410: инварианты", not probs, "; ".join(probs))


def scenario_dio_seven_of_eight():
    """Группа DIO, где 7 из 8 связок остаются целыми (фаза 1 дробления)."""
    rails = Rails([(1, "M", 400)] * 4)
    lens = [157.0]                      # модуль DIO64
    items, blocks = [0], [0]
    for b in range(1, 9):               # 8 связок: реле 110 + 2 барьера по 13
        it = list(range(len(lens), len(lens) + 3))
        lens += [110.0, 13.0, 13.0]
        items += it
        blocks += [b] * 3
    placed = {}
    before = rails.snapshot()
    budget = Budget()
    plan, hit = try_plan_blocks(rails, lens, 1, "M", items, blocks, budget)
    check("DIO 7/8: целыми блоками не укладывается", plan is None)
    check("DIO 7/8: неудача не меняет состояние", rails.snapshot() == before)
    res = try_place_group_in_cab(rails, lens, placed, 1, "M", items, blocks)
    check("DIO 7/8: группа размещена с дроблением", res["ok"] and res["was_split"])
    intact_b = intact_bundles(blocks, items, placed, skip={0})
    check("DIO 7/8: целы 7 из 8 связок", intact_b == 7, f"целы {intact_b}")
    probs = check_invariants("DIO", rails, lens, placed, [(items, 1)])
    check("DIO 7/8: инварианты", not probs, "; ".join(probs))


def scenario_bundle_wider_than_rail():
    """Связка шире рейки: размещения нет, состояние не тронуто."""
    rails = Rails([(1, "M", 200), (1, "M", 200)])
    lens = [250.0, 50.0]
    placed = {}
    before = rails.snapshot()
    ok = place_bundle_on_free_rail(rails, lens, placed, 1, "M", [0, 1])
    check("Шире рейки: целиком — отказ", not ok)
    res = try_place_group_in_cab(rails, lens, placed, 1, "M", [0, 1], [0, 0])
    check("Шире рейки: группой — отказ", not res["ok"] and not res["hit_limit"])
    check("Шире рейки: состояние не изменено",
          rails.snapshot() == before and not placed)


def scenario_tight_cabinet():
    """Тесный шкаф: частично занятые рейки, группа влезает только дроблением
    на элементы — и обратный случай, когда не влезает никак."""
    # Свободно 60/100/100; блоки 70, 70, 60, 60 (60 — составной 30+30):
    # целыми не встают, после дробления одного блока 60 — встают.
    rails = Rails([(1, "M", 100, 40), (1, "M", 100), (1, "M", 100)])
    lens = [70.0, 70.0, 30.0, 30.0, 60.0]
    items, blocks = list(range(5)), [0, 1, 2, 2, 3]
    placed = {}
    res = try_place_group_in_cab(rails, lens, placed, 1, "M", items, blocks)
    check("Тесный шкаф: размещено", res["ok"])
    probs = check_invariants("Тесный", rails, lens, placed, [(items, 1)])
    check("Тесный шкаф: инварианты", not probs, "; ".join(probs))

    rails2 = Rails([(1, "M", 100, 95), (1, "M", 100, 95)])
    placed2 = {}
    before = rails2.snapshot()
    res2 = try_place_group_in_cab(rails2, lens, placed2, 1, "M",
                                  [0, 1, 2], [0, 0, 1])
    check("Тесный шкаф: не влезает — отказ без hit_limit",
          not res2["ok"] and not res2["hit_limit"])
    check("Тесный шкаф: состояние не изменено", rails2.snapshot() == before)


def scenario_phase2_no_repeat():
    """Фаза 2 не повторяет первый вариант фазы 1.

    Блоки 110 (70+40), 80 (50+30) и 10 на двух рейках по 100: целыми не
    встают, фаза 1 дробит по очереди каждый составной блок, фаза 2 раньше
    начиналась с того же самого широкого блока — тот же набор, тот же
    отказ, но узлы сквозного бюджета потрачены. Итог укладки при этом не
    менялся: группу кладёт фоллбэк на отдельные элементы.
    """
    rails = Rails([(1, "M", 100), (1, "M", 100)])
    lens = [70.0, 40.0, 50.0, 30.0, 10.0]
    items, blocks = list(range(5)), [0, 0, 1, 1, 2]
    placed, budget = {}, Budget()
    res = try_place_group_in_cab(rails, lens, placed, 1, "M", items, blocks,
                                 budget=budget)
    check("Фаза 2: группа размещена", res["ok"])
    dup = len(budget.variants) - len(set(budget.variants))
    check("Фаза 2: один и тот же набор блоков не пробуется дважды", dup == 0,
          f"повторов {dup}: {budget.variants}")
    probs = check_invariants("Фаза 2", rails, lens, placed, [(items, 1)])
    check("Фаза 2: инварианты", not probs, "; ".join(probs))
    print(f"     (вызовов TryPlanBlocks {budget.plan_calls}, "
          f"узлов {budget.nodes})")


def scenario_budget_not_wasted():
    """Повтор фазы 2 не превращает решаемую задачу в ошибку 44/46.

    Задача с урезанным бюджетом (60 узлов): укладка существует, и её хватает
    бюджета — но прежде фаза 2 первым делом повторяла вариант фазы 1, поиск
    упирался в лимит, и вызывающий получал hitLimit («предел поиска»,
    ошибки 44/46) вместо укладки.
    """
    rails = Rails([(1, "M", 120), (1, "M", 120), (1, "M", 160), (1, "M", 200)])
    lens = [35.0, 67.0, 100.0, 100.0, 23.0, 67.0, 13.0, 100.0, 67.0]
    items, blocks = list(range(9)), [0, 0, 0, 1, 1, 2, 2, 3, 3]
    placed, budget = {}, Budget(limit=60)
    res = try_place_group_in_cab(rails, lens, placed, 1, "M", items, blocks,
                                 budget=budget)
    check("Бюджет не растрачен: решаемая задача уложена", res["ok"],
          f"узлов {budget.nodes} из {budget.limit}")
    check("Бюджет не растрачен: предел поиска не взведён", not res["hit_limit"])
    probs = check_invariants("Бюджет 60", rails, lens, placed, [(items, 1)])
    check("Бюджет не растрачен: инварианты", not probs, "; ".join(probs))


def scenario_shared_budget():
    """Сквозной бюджет: суммарно узлов не больше лимита + по одному
    «лишнему» узлу на вызов SearchBlocks, а hit_limit взведён."""
    # Детерминированно тяжёлая задача: без лимита зеркало тратит >200
    # узлов и укладывает группу дроблением; с лимитом 40 - hit_limit.
    rails = Rails([(1, "M", 120)] * 4)
    lens = [40.0, 50.0, 40.0, 40.0, 60.0, 50.0, 30.0,
            20.0, 30.0, 20.0, 20.0, 30.0, 30.0, 20.0]
    items = list(range(len(lens)))
    blocks = [i // 2 for i in items]
    budget = Budget(limit=40)
    res = try_place_group_in_cab(rails, lens, {}, 1, "M", items, blocks,
                                 budget=budget)
    check("Бюджет: поиск упёрся в лимит", res["hit_limit"])
    # Сам по себе «nodes > limit» ничего не доказывает: счётчик на вызов
    # тоже его превысит. Сквозной бюджет доказывает следующая проверка.
    check("Бюджет: лимит взведён", budget.nodes > budget.limit)
    check("Бюджет: счётчик сквозной — превышение не больше числа поисков",
          budget.nodes <= budget.limit + budget.search_calls,
          f"nodes={budget.nodes} calls={budget.search_calls}")


# ---------------------------------------------------------------------------
# Рандом со сверкой с брутфорсом.
# ---------------------------------------------------------------------------

def randomized(trials=600, seed=20261015, caps_pool=(80, 100, 120),
               used_pool=(0, 0, 20), widths=(10, 20, 30, 40, 50),
               max_blk=6, max_per_blk=3, label=""):
    """Рандом со сверкой с брутфорсом.

    Ширины по умолчанию круглые и кратны 10 — так часто выходит равное
    свободное место (проверка дедупликации реек в SearchBlocks). Второй
    прогон в main() берёт настоящие ширины модулей (13, 23, 157, …), где
    равных остатков почти не бывает.
    """
    rng = random.Random(seed)
    tag = f"{label} " if label else ""
    mism_plan = mism_group = state_dirty = inv_fail = 0
    skipped_limit = 0
    placed_ok = placed_no = 0
    for t in range(trials):
        n_rails = rng.randint(1, 4)
        caps = [float(rng.choice(caps_pool)) for _ in range(n_rails)]
        used = [float(rng.choice(used_pool)) for _ in range(n_rails)]
        n_blk = rng.randint(1, max_blk)
        lens, items, blocks = [], [], []
        for b in range(n_blk):
            for _ in range(rng.randint(1, max_per_blk)):
                items.append(len(lens))
                blocks.append(b)
                lens.append(float(rng.choice(widths)))
        free = [c - u for c, u in zip(caps, used)]

        # Целые блоки против брутфорса.
        blk_len = [0.0] * n_blk
        for idx, b in zip(items, blocks):
            blk_len[b] += lens[idx]
        rails = Rails([(1, "M", c, u) for c, u in zip(caps, used)])
        before = rails.snapshot()
        plan, hit = try_plan_blocks(rails, lens, 1, "M", items, blocks,
                                    Budget())
        if hit:
            skipped_limit += 1
        elif (plan is not None) != brute_force(blk_len, free):
            mism_plan += 1
        if rails.snapshot() != before:
            state_dirty += 1

        # Вся группа (с дроблением) против брутфорса по элементам.
        rails = Rails([(1, "M", c, u) for c, u in zip(caps, used)])
        before = rails.snapshot()
        placed = {}
        res = try_place_group_in_cab(rails, lens, placed, 1, "M",
                                     items, blocks)
        if res["hit_limit"] and not res["ok"]:
            skipped_limit += 1
        elif res["ok"] != brute_force(lens, free):
            mism_group += 1
        if res["ok"]:
            placed_ok += 1
            probs = check_invariants(f"rand{t}", rails, lens, placed,
                                     [(items, 1)])
            if probs:
                inv_fail += 1
        else:
            placed_no += 1
            if rails.snapshot() != before or placed:
                state_dirty += 1
    check(f"{tag}Рандом ({trials}, seed={seed}): целые блоки == брутфорс",
          mism_plan == 0, f"расхождений {mism_plan}")
    check(f"{tag}Рандом: группа с дроблением == брутфорс по элементам",
          mism_group == 0, f"расхождений {mism_group}")
    check(f"{tag}Рандом: неудача не пачкает состояние", state_dirty == 0,
          f"случаев {state_dirty}")
    check(f"{tag}Рандом: инварианты при успехе", inv_fail == 0,
          f"случаев {inv_fail}")
    # Прогон, в котором каждая задача упёрлась в лимит перебора, зелёный
    # пустой проверкой: сравнивать с брутфорсом было нечего. Задачи здесь
    # малы, лимит недостижим — если он всё-таки взведён, чинить генератор.
    check(f"{tag}Рандом: лимит перебора не достигнут", skipped_limit == 0,
          f"пропущено {skipped_limit}")
    # Второй способ выродиться — все задачи одного исхода.
    check(f"{tag}Рандом: есть и укладки, и отказы",
          placed_ok > 0 and placed_no > 0,
          f"уложено {placed_ok}, отказов {placed_no}")
    print(f"     (уложено {placed_ok}, отказов {placed_no}, "
          f"пропущено из-за лимита перебора: {skipped_limit})")


def main():
    scenario_p410_bundles()
    scenario_dio_seven_of_eight()
    scenario_bundle_wider_than_rail()
    scenario_tight_cabinet()
    scenario_phase2_no_repeat()
    scenario_budget_not_wasted()
    scenario_shared_budget()
    # Время прогона печатается: рандом сверяется с брутфорсом, а тот на
    # невозможных задачах экспоненциален. Дешёвый отказ в brute_force
    # снимает часть таких случаев, но не все (бывает, что сумма влезает,
    # а укладки всё равно нет) — рост времени должен быть виден в логе CI,
    # а не обнаруживаться через полгода.
    t0 = time.perf_counter()
    randomized()
    t1 = time.perf_counter()
    print(f"     (время прогона {t1 - t0:.2f} s)")
    # Второй прогон: настоящие ширины модулей и барьеров (defLensFallback),
    # до 8 блоков — остатки реек почти никогда не равны, дедупликация
    # SearchBlocks не помогает и перебор идёт шире. Ёмкости подобраны так,
    # чтобы задачи были пограничными: и целые блоки не влезают, и дробление
    # срабатывает (иначе прогон вырождается в одни отказы или одни укладки).
    randomized(trials=400, seed=20261103, caps_pool=(200, 260, 320),
               used_pool=(0, 0, 23), widths=(13, 23, 35, 67, 100, 110, 157, 164),
               max_blk=8, max_per_blk=2, label="Неровные ширины:")
    print(f"     (время прогона {time.perf_counter() - t1:.2f} s)")
    if FAIL:
        print(f"\nрасхождений: {len(FAIL)}")
        return 1
    print("\nplan-spec: все сценарии и рандомизированные прогоны сошлись")
    return 0


if __name__ == "__main__":
    # Вывод всегда в UTF-8: на Windows в pipe/файл (Git Bash, «> log», CI)
    # Python иначе пишет в ANSI-кодировке (cp1251/cp1252) и падает на «↔»,
    # «−», «—», а в cp1252 и на кириллице (UnicodeEncodeError). Сверяет
    # tools/check-sync.py запуском с PYTHONIOENCODING=cp1252.
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    sys.exit(main())
