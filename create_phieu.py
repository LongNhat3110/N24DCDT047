"""
Tạo file Word Phiếu thực hành CTDL-GT buổi 05
Sinh viên: Từ Lê Nhật Long - N24DCDT047 - D24CQDT01-N
"""
import zipfile, os, sys

output_path = r"c:\Users\admin\Downloads\NhatLong\extracted_zip\TuLeNhatLong_thuchanhbuoi05.docx"

# ===========================================================================
# NỘI DUNG CODE JULIA hoàn chỉnh 4 hàm
# ===========================================================================

CODE_MERGE_SORT = '''function merge_sort(a::AbstractVector, before::Function)
    n = length(a)
    # Điều kiện dừng: mảng rỗng hoặc một phần tử
    if n <= 1
        return collect(a)
    end
    # Chia đôi mảng
    mid = div(n, 2)
    left  = merge_sort(a[1:mid],   before)
    right = merge_sort(a[mid+1:n], before)
    # Trộn hai nửa đã có thứ tự
    result = similar(a, n)
    i = 1; j = 1; k = 1
    while i <= length(left) && j <= length(right)
        if before(right[j], left[i])
            result[k] = right[j]; j += 1
        else
            result[k] = left[i]; i += 1
        end
        k += 1
    end
    while i <= length(left);  result[k] = left[i];  i += 1; k += 1; end
    while j <= length(right); result[k] = right[j]; j += 1; k += 1; end
    result
end'''

CODE_BINARY_SEARCH = '''function binary_search(a::Vector{Order}, id::Int)
    left = 1
    right = length(a)
    while left <= right
        mid = left + div(right - left, 2)
        if a[mid].id == id
            return a[mid]        # tìm thấy
        elseif a[mid].id < id
            left = mid + 1       # id nằm nửa phải
        else
            right = mid - 1      # id nằm nửa trái
        end
    end
    nothing  # không tìm thấy
end'''

CODE_BFS = '''function bfs_routes(graph, source=1)
    n = length(graph)
    distance = fill(-1, n)
    parent   = zeros(Int, n)
    distance[source] = 0
    q = CircularQueue(n)
    enqueue!(q, source)
    while q.count > 0
        u = dequeue!(q)
        for v in graph[u]
            if distance[v] == -1
                distance[v] = distance[u] + 1
                parent[v]   = u
                enqueue!(q, v)
            end
        end
    end
    # Dựng routes: truy vết ngược từ mỗi đỉnh về nguồn
    routes = Vector{Vector{Int}}(undef, n)
    for v in 1:n
        if distance[v] == -1
            routes[v] = Int[]
        else
            path = Int[]; curr = v
            while curr != 0
                push!(path, curr); curr = parent[curr]
            end
            routes[v] = reverse(path)
        end
    end
    (distance=distance, parent=parent, routes=routes)
end'''

CODE_KNAPSACK = '''function knapsack(orders, costs, budget)
    budget >= 0 || throw(ArgumentError("Ngân sách phải không âm."))
    n = length(orders)
    # table[i+1, b+1] = F(i, b) = lợi ích tốt nhất của i đơn đầu / ngân sách b
    table = zeros(Int, n+1, budget+1)
    for i in 1:n
        o  = orders[i]
        τi = costs[o.id]
        for b in 0:budget
            not_take = table[i, b+1]
            take = (τi <= b) ? (o.benefit + table[i, b-τi+1]) : -1
            # Khi bằng: giữ phương án không lấy
            table[i+1, b+1] = (take > not_take) ? take : not_take
        end
    end
    # Truy vết
    ids = Int[]; b = budget
    for i in n:-1:1
        o  = orders[i]; τi = costs[o.id]
        if table[i+1, b+1] != table[i, b+1] && τi <= b
            push!(ids, o.id); b -= τi
        end
    end
    value = sum(o.benefit for o in orders if o.id in Set(ids); init=0)
    cost  = sum(costs[id] for id in ids; init=0)
    (ids=ids, value=value, cost=cost, table=table)
end'''

# ===========================================================================
# XML content_types
# ===========================================================================
CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
    <w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:numPr><w:ilvl w:val="0"/></w:numPr></w:pPr>
    <w:rPr><w:b/><w:sz w:val="32"/><w:color w:val="1F3864"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="2E75B6"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:rPr><w:b/><w:sz w:val="24"/><w:color w:val="2E75B6"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="18"/></w:rPr>
  </w:style>
</w:styles>'''

SETTINGS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="709"/>
</w:settings>'''

# ===========================================================================
# Helper functions to build XML paragraphs
# ===========================================================================
def esc(text):
    """Escape XML special characters"""
    return (text.replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;'))

def para(text, style="Normal", bold=False, size=None, color=None, indent=None):
    """Build a paragraph XML element"""
    ppr = ""
    if style != "Normal":
        ppr += f'<w:pStyle w:val="{style}"/>'
    if indent:
        ppr += f'<w:ind w:left="{indent}"/>'
    if ppr:
        ppr = f'<w:pPr>{ppr}</w:pPr>'
    
    rpr = ""
    if bold:
        rpr += "<w:b/>"
    if size:
        rpr += f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    if color:
        rpr += f'<w:color w:val="{color}"/>'
    if rpr:
        rpr = f'<w:rPr>{rpr}</w:rPr>'
    
    return f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r></w:p>'

def heading(text, level=1):
    colors = {1: "1F3864", 2: "2E75B6", 3: "2E74B5"}
    sizes  = {1: "32", 2: "28", 3: "24"}
    color = colors.get(level, "000000")
    size  = sizes.get(level, "24")
    return f'''<w:p>
      <w:pPr><w:pStyle w:val="Heading{level}"/></w:pPr>
      <w:r><w:rPr><w:b/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/><w:color w:val="{color}"/></w:rPr>
        <w:t>{esc(text)}</w:t></w:r></w:p>'''

def code_para(line):
    """Paragraph in monospace font for code"""
    return f'''<w:p>
      <w:pPr><w:pStyle w:val="Code"/><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/></w:pPr>
      <w:r><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
        <w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>
        <w:t xml:space="preserve">{esc(line)}</w:t></w:r></w:p>'''

def code_block(code_text):
    """Convert multi-line code string to docx XML paragraphs"""
    lines = code_text.split('\n')
    return '\n'.join(code_para(l) for l in lines)

def table_row(*cells, header=False):
    """Build a table row"""
    row_cells = ""
    for cell in cells:
        shd = '<w:shd w:val="clear" w:color="auto" w:fill="BDD7EE"/>' if header else ''
        bold = '<w:b/>' if header else ''
        row_cells += f'''<w:tc>
          <w:tcPr><w:tcBorders>
            <w:top w:val="single" w:sz="4" w:color="2E75B6"/>
            <w:bottom w:val="single" w:sz="4" w:color="2E75B6"/>
            <w:left w:val="single" w:sz="4" w:color="2E75B6"/>
            <w:right w:val="single" w:sz="4" w:color="2E75B6"/>
          </w:tcBorders>{shd}</w:tcPr>
          <w:p><w:r><w:rPr>{bold}</w:rPr><w:t xml:space="preserve">{esc(str(cell))}</w:t></w:r></w:p>
        </w:tc>'''
    return f'<w:tr>{row_cells}</w:tr>'

def table(*rows):
    """Build a full table"""
    inner = '\n'.join(rows)
    return f'''<w:tbl>
      <w:tblPr>
        <w:tblStyle w:val="TableGrid"/>
        <w:tblW w:w="9360" w:type="dxa"/>
        <w:tblBorders>
          <w:top w:val="single" w:sz="4" w:color="2E75B6"/>
          <w:bottom w:val="single" w:sz="4" w:color="2E75B6"/>
          <w:left w:val="single" w:sz="4" w:color="2E75B6"/>
          <w:right w:val="single" w:sz="4" w:color="2E75B6"/>
          <w:insideH w:val="single" w:sz="4" w:color="2E75B6"/>
          <w:insideV w:val="single" w:sz="4" w:color="2E75B6"/>
        </w:tblBorders>
      </w:tblPr>
      {inner}
    </w:tbl>'''

# ===========================================================================
# Build full document XML
# ===========================================================================
paras = []

# ---- Tiêu đề ----
paras.append(para("PHIẾU THỰC HÀNH — 09/10/2026", bold=True, size="40", color="1F3864"))
paras.append(para("Môn học: Cấu trúc dữ liệu và Giải thuật", bold=True, size="24"))
paras.append(para(""))

# ---- Thông tin sinh viên ----
paras.append(table(
    table_row("Họ tên", "Từ Lê Nhật Long", "MSSV", "N24DCDT047", header=True),
    table_row("Lớp", "D24CQDT01-N", "Ngày thực hành", "09/10/2026"),
    table_row("GitHub", "https://github.com/LongNhat3110/N24DCDT047", "Commit SHA", "6eec32d54b8fc866021fd957b4c378c83ece4491"),
))
paras.append(para(""))

# ============================
# MỤC 1: Mô tả một luồng ứng dụng
# ============================
paras.append(heading("1. Mô tả một luồng ứng dụng", 1))
paras.append(para("Mục tiêu: Điều phối giao đơn trong khuôn viên trường, chọn tập đơn tối ưu trong ngân sách thời gian."))
paras.append(para(""))
paras.append(para("Đầu vào:", bold=True))
paras.append(para("  • don_hang.csv: danh sách 3 đơn (id, diem_giao, thoi_gian_phuc_vu, loi_ich)"))
paras.append(para("  • ban_do.csv: 7 cạnh của đồ thị vô hướng 6 đỉnh"))
paras.append(para("  • Ngân sách B = 8 (bộ MINI001)"))
paras.append(para(""))
paras.append(para("Đầu ra:", bold=True))
paras.append(para("  • Tập đơn được chọn (IDs), tổng thời gian, tổng lợi ích"))
paras.append(para("  • ket_qua.csv: trạng thái từng đơn (9 cột)"))
paras.append(para(""))
paras.append(para("Sơ đồ luồng dữ liệu:", bold=True))
paras.append(para("  CSV → read_data() → orders + edges"))
paras.append(para("  edges → adjacency() → graph (danh sách kề)"))
paras.append(para("  graph → [NV3] bfs_routes() → distance, parent, routes"))
paras.append(para("  distance + orders → costs = Dict(id → τ = 2h + service)"))
paras.append(para("  orders → [NV1] merge_sort() → sorted_ids (theo ID tăng)"))
paras.append(para("  sorted_ids → [NV2] binary_search() → tra cứu đơn theo ID"))
paras.append(para("  sorted_ids + costs + budget → [NV4] knapsack() → selected_ids"))
paras.append(para("  selected_ids → priority_orders() → FIFO hàng đợi → deliver!/undo!"))
paras.append(para("  → export_result() → ket_qua.csv"))
paras.append(para(""))

# ============================
# MỤC 2: Bảng truy vết đường và thời gian
# ============================
paras.append(heading("2. Bảng truy vết đường và thời gian", 1))
paras.append(para("Bản đồ: 6 đỉnh, 7 cạnh: (1,2),(1,3),(2,4),(3,4),(3,5),(4,6),(5,6)"))
paras.append(para("BFS từ nguồn 1: khoảng cách = [0, 1, 1, 2, 2, 3]"))
paras.append(para(""))
paras.append(table(
    table_row("ID Đơn", "Điểm giao", "Đường lượt đi", "Số cạnh h", "2h + phục vụ (τ)", "Lợi ích", header=True),
    table_row("1 (A)", "2", "1 → 2", "1", "2×1 + 1 = 3", "6"),
    table_row("2 (B)", "4", "1 → 2 → 4", "2", "2×2 + 1 = 5", "8"),
    table_row("3 (C)", "6", "1 → 2 → 4 → 6", "3", "2×3 + 1 = 7", "10"),
))
paras.append(para(""))
paras.append(para("Chi phí (costs): Dict(1 => 3, 2 => 5, 3 => 7)"))
paras.append(para("Tỷ số ưu tiên: đơn 1: 6/3=2.0; đơn 2: 8/5=1.6; đơn 3: 10/7≈1.43 → thứ tự: [1, 2, 3]"))
paras.append(para(""))

# ============================
# MỤC 3: Minh chứng bốn hàm
# ============================
paras.append(heading("3. Minh chứng bốn hàm", 1))

# NV1
paras.append(heading("NV1 · Hoàn thiện merge_sort(a, before)", 2))
paras.append(para("Ý tưởng / điều kiện áp dụng:", bold=True))
paras.append(para("Sắp xếp trộn (Merge Sort) dùng phương pháp chia để trị:"))
paras.append(para("  (1) Điều kiện dừng: length(a) ≤ 1 → trả bản sao của a."))
paras.append(para("  (2) Chia: mid = div(n, 2); đệ quy sắp left = a[1:mid], right = a[mid+1:n]."))
paras.append(para("  (3) Trộn: so sánh hai đầu nhánh bằng before(right[j], left[i]):"))
paras.append(para("      • Nếu đúng: lấy right[j] trước."))
paras.append(para("      • Ngược lại (kể cả hòa): lấy left[i] → đảm bảo tính ổn định (stable)."))
paras.append(para("  (4) Chép phần dư của nhánh còn lại."))
paras.append(para("  Không gọi sort/sort! có sẵn. Không làm thay đổi mảng đầu vào."))
paras.append(para(""))
paras.append(para("Code hoàn chỉnh:", bold=True))
paras.append(code_block(CODE_MERGE_SORT))
paras.append(para(""))
paras.append(para("Lời gọi và kết quả kiểm thử:", bold=True))
paras.append(para("  julia TH_2026_10_09_Kiem_thu.jl --phase=sort"))
paras.append(para("  Kết quả: DAT pha sort: 8 phép kiểm tra. ✅"))
paras.append(para("  • merge_sort(Int[], (<)) == Int[]                              → ✅ dãy rỗng"))
paras.append(para("  • merge_sort([7], (<)) == [7]                                  → ✅ một phần tử"))
paras.append(para("  • input=[7,2,5,1]; output=[1,2,5,7]; input không đổi          → ✅"))
paras.append(para("  • merge_sort([2,1,2,3], (>)) == [3,2,2,1]                     → ✅ giảm dần"))
paras.append(para("  • Sắp ổn định: [(1,\"B\"),(1,\"D\"),(2,\"A\"),(2,\"C\")]            → ✅"))
paras.append(para("  • Sắp Order [3,1,2] → IDs [1,2,3]                             → ✅"))
paras.append(para("  • Tỷ số hòa với tie=0: ID tăng; tie=1: ID giảm               → ✅"))
paras.append(para(""))
paras.append(para("Chi phí thời gian / bộ nhớ:", bold=True))
paras.append(para("  • Thời gian: O(n log n) — có log₂n tầng chia, mỗi tầng trộn O(n) phần tử."))
paras.append(para("  • Bộ nhớ phụ: O(n) cho kết quả và các lát cắt tạm, O(log n) stack đệ quy."))
paras.append(para(""))

# NV2
paras.append(heading("NV2 · Hoàn thiện binary_search(a::Vector{Order}, id::Int)", 2))
paras.append(para("Ý tưởng / điều kiện áp dụng:", bold=True))
paras.append(para("Tiền điều kiện: mảng a phải sắp tăng theo Order.id (đảm bảo bởi merge_sort ở NV1)."))
paras.append(para("  (1) Duy trì đoạn [left, right] là vùng còn khả năng chứa id."))
paras.append(para("  (2) Mỗi bước: mid = left + div(right-left, 2)"))
paras.append(para("      • a[mid].id == id → trả a[mid]"))
paras.append(para("      • a[mid].id < id  → left = mid + 1  (loại nửa trái)"))
paras.append(para("      • a[mid].id > id  → right = mid - 1 (loại nửa phải)"))
paras.append(para("  (3) left > right: đoạn rỗng → trả nothing."))
paras.append(para("  Chỉ truy cập a[mid] trong vòng lặp left ≤ right → an toàn với mảng rỗng."))
paras.append(para(""))
paras.append(para("Code hoàn chỉnh:", bold=True))
paras.append(code_block(CODE_BINARY_SEARCH))
paras.append(para(""))
paras.append(para("Bảng truy vết tìm ID=2 trên dãy [1, 2, 3]:", bold=True))
paras.append(table(
    table_row("Lần lặp", "left", "right", "mid", "a[mid].id", "Quyết định", header=True),
    table_row("1", "1", "3", "2", "2", "Bằng → trả a[2]"),
))
paras.append(para(""))
paras.append(para("Bảng truy vết tìm ID=99 trên dãy [1, 2, 3]:", bold=True))
paras.append(table(
    table_row("Lần lặp", "left", "right", "mid", "a[mid].id", "Quyết định", header=True),
    table_row("1", "1", "3", "2", "2", "2 < 99 → left = 3"),
    table_row("2", "3", "3", "3", "3", "3 < 99 → left = 4"),
    table_row("3", "4", "3", "-", "-", "left > right → nothing"),
))
paras.append(para(""))
paras.append(para("Lời gọi và kết quả kiểm thử:", bold=True))
paras.append(para("  julia TH_2026_10_09_Kiem_thu.jl --phase=search"))
paras.append(para("  Kết quả: DAT pha search: 5 phép kiểm tra. ✅"))
paras.append(para("  • binary_search(Order[], 2) === nothing                        → ✅ mảng rỗng"))
paras.append(para("  • binary_search(a, 2).id == 2                                  → ✅ ID đầu dãy"))
paras.append(para("  • binary_search(a, 7).id == 7                                  → ✅ ID cuối dãy"))
paras.append(para("  • binary_search(a, 4).benefit == 8 (trả Order, không trả vị trí) → ✅"))
paras.append(para("  • binary_search(a, 3) === nothing && binary_search(a,99)===nothing → ✅"))
paras.append(para(""))
paras.append(para("Chi phí thời gian / bộ nhớ:", bold=True))
paras.append(para("  • Thời gian: O(log n) — mỗi lần lặp giảm đoạn xét còn một nửa."))
paras.append(para("  • Bộ nhớ phụ: O(1) — chỉ dùng 3 biến left, right, mid."))
paras.append(para("  So sánh với BST: BST tra cứu O(h), tốt nhất O(log n) nếu cân bằng, xấu nhất O(n)."))
paras.append(para(""))

# NV3
paras.append(heading("NV3 · Hoàn thiện bfs_routes(graph, source=1)", 2))
paras.append(para("Ý tưởng / điều kiện áp dụng:", bold=True))
paras.append(para("BFS (Breadth-First Search) trên đồ thị không trọng số:"))
paras.append(para("  (1) Khởi tạo: distance = [-1,...,-1]; distance[source] = 0; đưa source vào FIFO."))
paras.append(para("  (2) Duyệt FIFO: lấy đỉnh u → duyệt láng giềng v:"))
paras.append(para("      • Nếu distance[v] == -1 (chưa thăm): ghi distance[v] = distance[u]+1,"))
paras.append(para("        parent[v] = u, đưa v vào FIFO (đánh dấu KHI ĐƯA VÀO, không khi lấy ra)."))
paras.append(para("  (3) Sau BFS: dựng routes[v] = đảo ngược đường ngược từ v theo parent về source."))
paras.append(para("  (4) Đỉnh không tới được: distance = -1, routes = []."))
paras.append(para(""))
paras.append(para("Truy vết BFS bộ MINI001 (láng giềng tăng dần theo tie=0):", bold=True))
paras.append(table(
    table_row("Bước", "Xử lý u", "Phát hiện v", "distance[v]", "parent[v]", "FIFO sau bước", header=True),
    table_row("0", "-", "source=1", "d[1]=0", "-", "[1]"),
    table_row("1", "1", "2, 3", "d[2]=1, d[3]=1", "p[2]=1, p[3]=1", "[2, 3]"),
    table_row("2", "2", "4", "d[4]=2", "p[4]=2", "[3, 4]"),
    table_row("3", "3", "5", "d[5]=2", "p[5]=3", "[4, 5]"),
    table_row("4", "4", "6", "d[6]=3", "p[6]=4", "[5, 6]"),
    table_row("5", "5", "-", "-", "-", "[6]"),
    table_row("6", "6", "-", "-", "-", "[]"),
))
paras.append(para("  Kết quả: distance = [0, 1, 1, 2, 2, 3] ✅"))
paras.append(para(""))
paras.append(para("Code hoàn chỉnh:", bold=True))
paras.append(code_block(CODE_BFS))
paras.append(para(""))
paras.append(para("Lời gọi và kết quả kiểm thử:", bold=True))
paras.append(para("  julia TH_2026_10_09_Kiem_thu.jl --phase=bfs"))
paras.append(para("  Kết quả: DAT pha bfs: 7 phép kiểm tra. ✅"))
paras.append(para("  • distance == [0,1,1,2,2,3]                                    → ✅"))
paras.append(para("  • Mỗi tuyến hợp lệ và khớp chuỗi cha/khoảng cách              → ✅"))
paras.append(para("  • routes[1] == [1] && distance[1] == 0                         → ✅ nguồn"))
paras.append(para("  • [2*d[o.destination]+o.service for o in orders] == [3,5,7]    → ✅"))
paras.append(para("  • Điểm không tới: distance=-1, routes=[]                       → ✅"))
paras.append(para("  • BFS từ nguồn khác (đỉnh 4): đúng khoảng cách và tuyến       → ✅"))
paras.append(para(""))
paras.append(para("Chi phí thời gian / bộ nhớ:", bold=True))
paras.append(para("  • Thời gian: O(V + E) — mỗi đỉnh và cạnh được xét đúng một lần."))
paras.append(para("  • Bộ nhớ: O(V) cho distance, parent, FIFO; O(V²) cho routes nếu lưu tất cả."))
paras.append(para("  • Tính costs cho n đơn: O(n) — không cần chạy lại BFS từng đơn."))
paras.append(para(""))

# NV4
paras.append(heading("NV4 · Hoàn thiện knapsack(orders, costs, budget)", 2))
paras.append(para("Ý tưởng / điều kiện áp dụng:", bold=True))
paras.append(para("Bài toán ba lô 0/1: chọn tối đa mỗi đơn một lần để tối đa lợi ích trong ngân sách B."))
paras.append(para("  Trạng thái: F(i, b) = lợi ích tốt nhất khi xét i đơn đầu, ngân sách b."))
paras.append(para("  Công thức:"))
paras.append(para("      F(0, b) = 0 với mọi b (hàng cơ sở)"))
paras.append(para("      F(i, b) = F(i-1, b)                               nếu τ_i > b"))
paras.append(para("      F(i, b) = max(F(i-1, b),  benefit_i + F(i-1, b-τ_i))   nếu τ_i ≤ b"))
paras.append(para("  Trong Julia: table[i+1, b+1] = F(i, b) (chỉ số từ 1)."))
paras.append(para("  Truy vết: i=n..1 xét ngược; nếu table[i+1,b+1] ≠ table[i,b+1] → đã chọn đơn i."))
paras.append(para(""))
paras.append(para("Điền bảng DP bộ MINI001 (n=3, B=8):", bold=True))
paras.append(para("  Đơn 1: τ=3, lợi ích=6  |  Đơn 2: τ=5, lợi ích=8  |  Đơn 3: τ=7, lợi ích=10"))
paras.append(para(""))
paras.append(table(
    table_row("F(i,b)", "b=0", "b=1", "b=2", "b=3", "b=4", "b=5", "b=6", "b=7", "b=8", header=True),
    table_row("i=0", "0", "0", "0", "0", "0", "0", "0", "0", "0"),
    table_row("i=1 (τ=3,v=6)", "0", "0", "0", "6", "6", "6", "6", "6", "6"),
    table_row("i=2 (τ=5,v=8)", "0", "0", "0", "6", "6", "8", "8", "8", "14"),
    table_row("i=3 (τ=7,v=10)", "0", "0", "0", "6", "6", "8", "8", "8", "14"),
))
paras.append(para(""))
paras.append(para("Giải thích ô F(2,8) = max(F(1,8), 8+F(1,3)) = max(6, 8+6) = 14 → chọn đơn 2.", bold=True))
paras.append(para("Giải thích ô F(3,8) = max(F(2,8), 10+F(2,1)) = max(14, 10+0) = 14 → KHÔNG chọn đơn 3."))
paras.append(para(""))
paras.append(para("Truy vết kết quả:", bold=True))
paras.append(para("  • i=3, b=8: F(3,8)=14 == F(2,8)=14 → KHÔNG chọn đơn 3, b giữ nguyên = 8"))
paras.append(para("  • i=2, b=8: F(2,8)=14 ≠ F(1,8)=6  → CHỌN đơn 2, b = 8-5 = 3"))
paras.append(para("  • i=1, b=3: F(1,3)=6  ≠ F(0,3)=0  → CHỌN đơn 1, b = 3-3 = 0"))
paras.append(para("  → Kết quả: ids=[1,2], cost=8, value=14 ✅"))
paras.append(para(""))
paras.append(para("Code hoàn chỉnh:", bold=True))
paras.append(code_block(CODE_KNAPSACK))
paras.append(para(""))
paras.append(para("Lời gọi và kết quả kiểm thử:", bold=True))
paras.append(para("  julia TH_2026_10_09_Kiem_thu.jl --phase=dp"))
paras.append(para("  Kết quả: DAT pha dp: 12 phép kiểm tra. ✅"))
paras.append(para("  • Set(ids)=={1,2}, value==14, cost==8                          → ✅"))
paras.append(para("  • B=0: tập rỗng, cost=0, value=0                              → ✅"))
paras.append(para("  • Một đơn duy nhất, B=8: ids=[1], value=6                     → ✅ không tái dùng"))
paras.append(para("  • Phản ví dụ tham lam: optimal.value=20 > greedy.value=19     → ✅ tham lam không tối ưu"))
paras.append(para("  • table[2,4]=6, table[3,9]=14                                 → ✅ ô DP đúng"))
paras.append(para("  • Đảo thứ tự đầu vào: vẫn cùng value và ids                  → ✅"))
paras.append(para(""))
paras.append(para("Chi phí thời gian / bộ nhớ:", bold=True))
paras.append(para("  • Thời gian: O(n × B) — điền bảng n hàng × B+1 cột."))
paras.append(para("  • Bộ nhớ: O(n × B) cho bảng DP + O(n) cho truy vết."))
paras.append(para("  • Vét cạn (đối chiếu): O(2^n × n) — chỉ dùng cho n nhỏ (≤ 20)."))
paras.append(para(""))

# ============================
# MỤC 4: Quyết định chọn đơn
# ============================
paras.append(heading("4. Quyết định chọn đơn", 1))
paras.append(para("Ngân sách: B = 8  |  Bộ dữ liệu: MINI001"))
paras.append(para(""))
paras.append(table(
    table_row("Tham số", "Giá trị", header=True),
    table_row("Ngân sách B", "8"),
    table_row("Tập đơn được chọn (DP)", "{1, 2}"),
    table_row("Tổng thời gian (cost)", "3 + 5 = 8"),
    table_row("Tổng lợi ích (value)", "6 + 8 = 14"),
    table_row("Kết quả vét cạn (oracle)", "14 ✅ (khớp DP)"),
    table_row("Kết quả tham lam (greedy)", "14 (bộ mini trùng với DP)"),
))
paras.append(para(""))
paras.append(para("Giải thích tại sao không chọn đơn 3 (τ=7, lợi ích=10):", bold=True))
paras.append(para("  • Nếu chọn đơn 3 (τ=7): ngân sách còn 1, không đủ lấy thêm đơn nào → lợi ích=10."))
paras.append(para("  • Nếu chọn đơn 1+2 (τ=3+5=8): lợi ích = 6+8 = 14 > 10 → tối ưu hơn."))
paras.append(para("  • DP chọn đúng tập {1,2} với lợi ích 14, xác nhận bằng vét cạn."))
paras.append(para(""))
paras.append(para("Ô F(2,8) = max(F(1,8), benefit_2 + F(1, 8-τ_2))"))
paras.append(para("         = max(6, 8 + F(1,3))"))
paras.append(para("         = max(6, 8 + 6)"))
paras.append(para("         = max(6, 14) = 14 → chọn đơn 2."))
paras.append(para("Vị trí trong Julia: table[3, 9] (i=2 → hàng 3; b=8 → cột 9)."))
paras.append(para(""))
paras.append(para("Tham lam theo tỷ số lợi ích/τ: [6/3=2.0, 8/5=1.6, 10/7=1.43] → chọn 1 rồi 2 (cùng tối ưu ở bộ mini)."))
paras.append(para("Phản ví dụ: orders [v=9,τ=4],[v=10,τ=5],[v=10,τ=5], B=10: tham lam chọn {1}+{2}=19 < DP chọn {2,3}=20."))
paras.append(para("→ Tham lam theo tỷ số KHÔNG bảo đảm tối ưu cho bài toán ba lô 0/1."))
paras.append(para(""))

# ============================
# MỤC 5: Giao và hoàn tác
# ============================
paras.append(heading("5. Giao và hoàn tác", 1))
paras.append(para("Sau khi lập kế hoạch: selected_ids = [1, 2], FIFO = [1, 2], stack = []."))
paras.append(para(""))
paras.append(table(
    table_row("Trạng thái", "FIFO (Queue)", "Stack (Lịch sử)", "Đã dùng (τ)", "Đã nhận (lợi ích)", header=True),
    table_row("Sau lập kế hoạch", "[1, 2]", "[]", "0", "0"),
    table_row("Sau giao đơn 1", "[2]", "[1]", "3", "6"),
    table_row("Sau hoàn tác đơn 1", "[1, 2]", "[]", "0", "0"),
    table_row("Giao lại đơn 1", "[2]", "[1]", "3", "6"),
    table_row("Giao tiếp đơn 2", "[]", "[1, 2]", "8", "14"),
))
paras.append(para(""))
paras.append(para("Kết quả chạy thực tế (--mini):", bold=True))
paras.append(code_block("""MINI001 | Ngân sách 8 | DP 14 | Vét cạn 14

MSSV: MINI001 | Ngân sách: 8
ID | Điểm giao | Phục vụ | Chi phí BFS | Lợi ích | Tuyến đi
1 | 2 | 1 | 3 | 6 | 1 -> 2
2 | 4 | 1 | 5 | 8 | 1 -> 2 -> 4
3 | 6 | 1 | 7 | 10 | 1 -> 2 -> 4 -> 6
Đã chọn: [1, 2] | Đang chờ: [1, 2]
Đã giao: Int64[] | Đã dùng: 0 | Lợi ích đã nhận: 0
Giao đơn 1 | Tuyến 1 -> 2 -> về kho 1 | Thời gian 3 | Lợi ích 6
Đang chờ [2] | Đã giao [1] | Đã dùng 3 | Đã nhận 6
Hoàn tác đơn 1 | Tuyến 1 -> 2 -> về kho 1 | Thời gian 3 | Lợi ích 6
Đang chờ [1, 2] | Đã giao Int64[] | Đã dùng 0 | Đã nhận 0
Giao đơn 1 ... | Đã dùng 3 | Đã nhận 6
Giao đơn 2 ... | Đã dùng 8 | Đã nhận 14
Đã xuất: ...\\data\\MINI001\\ket_qua.csv"""))
paras.append(para(""))
paras.append(para("Bất biến trạng thái kiểm tra không mất/lặp đơn:", bold=True))
paras.append(para("  Set(queue_values(pending) ∪ completed) == Set(selected_ids)"))
paras.append(para("  spent = Σ costs[id] cho id trong completed"))
paras.append(para("  earned = Σ benefit[id] cho id trong completed"))
paras.append(para("  0 ≤ spent ≤ Σcosts[selected] ≤ budget"))
paras.append(para(""))

# ============================
# MỤC 6: Đầu ra và nhận xét
# ============================
paras.append(heading("6. Đầu ra và nhận xét", 1))
paras.append(para("Đường dẫn CSV kết quả:", bold=True))
paras.append(para("  MINI001: extracted_zip\\data\\MINI001\\ket_qua.csv"))
paras.append(para("  N24DCDT047: extracted_zip\\data\\N24DCDT047\\ket_qua.csv"))
paras.append(para(""))
paras.append(para("Nội dung ket_qua.csv (MINI001):", bold=True))
paras.append(code_block("""id,diem_giao,so_canh,chi_phi,loi_ich,tuyen_di,duoc_chon,da_giao,trang_thai
1,2,1,3,6,1-2,1,1,da_giao
2,4,2,5,8,1-2-4,1,1,da_giao
3,6,3,7,10,1-2-4-6,0,0,khong_chon"""))
paras.append(para(""))
paras.append(para("Kết quả kiểm thử toàn bộ:", bold=True))
paras.append(table(
    table_row("Pha", "Lệnh chạy", "KQ mong đợi", "KQ thực tế", "Đạt?", header=True),
    table_row("data", "--phase=data", "12 tests", "12 tests", "✅"),
    table_row("sort", "--phase=sort", "8 tests", "8 tests", "✅"),
    table_row("search", "--phase=search", "5 tests", "5 tests", "✅"),
    table_row("bfs", "--phase=bfs", "7 tests", "7 tests", "✅"),
    table_row("dp", "--phase=dp", "12 tests", "12 tests", "✅"),
    table_row("app", "--phase=app", "16 tests", "16 tests", "✅"),
    table_row("all", "--phase=all", "60 tests", "60 tests", "✅"),
))
paras.append(para(""))
paras.append(para("Link notebook: TH_2026_10_09_Pluto.jl (chạy trực tiếp bằng Julia nếu chưa cài Pluto)"))
paras.append(para(""))
paras.append(para("Lỗi đã phát hiện và cách sửa:", bold=True))
paras.append(para("  Lỗi 1 (merge_sort): Ban đầu quên chép phần dư của nhánh còn lại sau khi một nhánh hết."))
paras.append(para("    → Sửa: Thêm hai vòng while độc lập để chép phần dư left và right."))
paras.append(para("  Lỗi 2 (knapsack): Dùng costs[i] (vị trí) thay vì costs[orders[i].id] → sai khi thứ tự đổi."))
paras.append(para("    → Sửa: Luôn tra bằng costs[o.id] với o = orders[i]."))
paras.append(para("  Lỗi 3 (bfs_routes): Quên đặt parent[source] = 0 ban đầu, gây vòng lặp khi truy vết."))
paras.append(para("    → Sửa: zeros(Int, n) khởi tạo parent = 0, vòng truy vết dừng khi curr == 0."))
paras.append(para(""))
paras.append(para("Một chức năng thay đổi ảnh hưởng đến chức năng sau:", bold=True))
paras.append(para("  • merge_sort là nền tảng cho NV2, NV3, NV4:"))
paras.append(para("    - NV2 (binary_search) yêu cầu mảng tăng theo ID — do merge_sort tạo."))
paras.append(para("    - NV3 (adjacency) sắp láng giềng bằng merge_sort → ảnh hưởng thứ tự duyệt BFS."))
paras.append(para("    - plan! dùng priority_orders (dùng merge_sort) để tạo FIFO theo ưu tiên."))
paras.append(para("  • bfs_routes tính costs → costs truyền vào knapsack → knapsack quyết định tập chọn."))
paras.append(para("  → Nếu bfs_routes sai → costs sai → knapsack chọn sai → toàn bộ kế hoạch sai."))
paras.append(para(""))

# ============================
# MỤC 7: Câu hỏi cuối buổi
# ============================
paras.append(heading("Câu hỏi cuối buổi", 1))

paras.append(heading("Q1.", 2))
paras.append(para("Vì sao không được dùng dãy đã sắp giảm theo lợi ích/τ để tìm kiếm nhị phân theo ID? Nêu tiền điều kiện của hàm tìm kiếm."))
paras.append(para(""))
paras.append(para("Trả lời:", bold=True))
paras.append(para("  Tìm kiếm nhị phân có tiền điều kiện: dãy phải sắp TĂNG theo KHÓA TÌM KIẾM."))
paras.append(para("  Khi ta cần tra theo ID, dãy phải tăng theo Order.id."))
paras.append(para("  Dãy sắp giảm theo lợi ích/τ có thứ tự khóa (ID) hoàn toàn ngẫu nhiên → tìm kiếm nhị phân"))
paras.append(para("  sẽ loại sai nửa mảng → không tìm thấy mục tiêu dù nó tồn tại trong dãy."))
paras.append(para("  Ví dụ: dãy ưu tiên = [ID1, ID2, ID3] (tỷ số 2.0, 1.6, 1.43) → thứ tự ID tăng dần."))
paras.append(para("  Nhưng nếu tỷ số khác nhau, thứ tự ưu tiên có thể là [ID3, ID1, ID2] → sắp giảm tỷ số,"))
paras.append(para("  không theo ID → binary_search theo ID sẽ sai hoàn toàn."))
paras.append(para(""))

paras.append(heading("Q2.", 2))
paras.append(para("Với đơn ID 2 tới điểm 4, phục vụ 1 và bản đồ mẫu, giải thích từng thành phần của thời gian thực hiện τ. BFS có tiếp tục cho thời gian ngắn nhất nếu các cạnh mang thời gian khác nhau không?"))
paras.append(para(""))
paras.append(para("Trả lời:", bold=True))
paras.append(para("  τ₂ = 2 × h₂ + s₂ = 2 × 2 + 1 = 5"))
paras.append(para("  Trong đó:"))
paras.append(para("    • h₂ = 2: số cạnh đường BFS ngắn nhất từ kho 1 → điểm 4 (đường 1→2→4)."))
paras.append(para("    • Nhân 2: vì phải đi VÀ VỀ (kho 1 → điểm 4 → kho 1) = 2 lượt × 2 cạnh."))
paras.append(para("    • s₂ = 1: thời gian phục vụ tại điểm 4 (bốc/dỡ hàng)."))
paras.append(para("  BFS chỉ đúng với đồ thị KHÔNG TRỌNG SỐ (mỗi cạnh tốn 1 đơn vị thời gian)."))
paras.append(para("  Nếu các cạnh có trọng số khác nhau, cần dùng thuật toán Dijkstra (không dùng BFS)."))
paras.append(para("  BFS tối thiểu hóa số cạnh, không phải tổng trọng số."))
paras.append(para(""))

paras.append(heading("Q3.", 2))
paras.append(para("Giải thích hai phương án tạo ô F(2,8) và vị trí của ô đó trong table Julia. Vì sao nhánh lấy đơn phải dùng hàng trước?"))
paras.append(para(""))
paras.append(para("Trả lời:", bold=True))
paras.append(para("  F(2,8) là lợi ích tốt nhất khi xét 2 đơn đầu (đơn 1 và 2) với ngân sách 8."))
paras.append(para("  Hai phương án:"))
paras.append(para("    • Không lấy đơn 2: F(2,8) = F(1,8) = 6 (chỉ dùng đơn 1)."))
paras.append(para("    • Lấy đơn 2 (τ₂=5): F(2,8) = benefit₂ + F(1, 8-5) = 8 + F(1,3) = 8+6 = 14."))
paras.append(para("  Chọn max(6, 14) = 14 → lấy đơn 2."))
paras.append(para("  Vị trí trong Julia: table[i+1, b+1] = table[2+1, 8+1] = table[3, 9]."))
paras.append(para("  Vì sao nhánh lấy phải dùng hàng trước (F(i-1, ...)):"))
paras.append(para("    Nếu dùng F(i, b-τ_i) thay vì F(i-1, b-τ_i) → bài toán ba lô không giới hạn (unbounded)"))
paras.append(para("    → cho phép lấy đơn i nhiều lần. Phải dùng hàng i-1 để đảm bảo mỗi đơn chỉ chọn một lần."))
paras.append(para(""))

paras.append(heading("Q4.", 2))
paras.append(para("Sau khi giao ID 1 rồi hoàn tác ngay, FIFO, ngăn xếp, tổng thời gian và tổng lợi ích phải thay đổi thế nào? Nêu một bất biến kiểm tra việc không mất hoặc lặp đơn."))
paras.append(para(""))
paras.append(para("Trả lời:", bold=True))
paras.append(para("  Trước khi giao:   FIFO=[1,2], stack=[], spent=0, earned=0"))
paras.append(para("  Sau giao ID 1:    FIFO=[2],   stack=[1], spent=3, earned=6"))
paras.append(para("  Sau hoàn tác ID1: FIFO=[1,2], stack=[],  spent=0, earned=0"))
paras.append(para("  → Trạng thái phục hồi hoàn toàn về ban đầu. ID 1 được đưa lại vào ĐẦU FIFO (không phải cuối)."))
paras.append(para(""))
paras.append(para("  Bất biến không mất / không lặp đơn:"))
paras.append(para("    Set(queue_values(pending) ∪ completed) == Set(selected_ids)"))
paras.append(para("    length(Set(pending ∪ completed)) == length(pending) + length(completed)"))
paras.append(para("  Nghĩa là: hợp của FIFO và lịch sử giao phải đúng bằng tập đơn chọn, không mất và không trùng."))
paras.append(para(""))

paras.append(heading("Q5.", 2))
paras.append(para("Cho biết độ phức tạp theo n,V,E,B của sắp xếp trộn, tìm kiếm nhị phân, BFS cốt lõi và DP bảng đầy đủ. Vì sao không thể khẳng định mọi BST đều tra cứu trong O(log n)?"))
paras.append(para(""))
paras.append(para("Trả lời:", bold=True))
paras.append(table(
    table_row("Thuật toán", "Thời gian", "Bộ nhớ phụ", header=True),
    table_row("Sắp xếp trộn", "O(n log n)", "O(n) buffer + O(log n) stack"),
    table_row("Tìm kiếm nhị phân", "O(log n)", "O(1)"),
    table_row("BFS cốt lõi", "O(V + E)", "O(V) (distance, parent, FIFO)"),
    table_row("Knapsack DP (bảng đầy đủ)", "O(n × B)", "O(n × B) + O(n) truy vết"),
))
paras.append(para(""))
paras.append(para("  BST không đảm bảo O(log n) vì:"))
paras.append(para("    • O(log n) chỉ đúng khi BST cân bằng (h ≈ log n)."))
paras.append(para("    • BST cơ bản (không tự cân bằng): nếu chèn dãy ID tăng dần [1,2,3,...,n]"))
paras.append(para("      → cây suy biến thành danh sách liên kết, h = n → tra cứu O(n)."))
paras.append(para("    • Ví dụ: chèn [1,2,3]: gốc=1, phải=2, phải phải=3 → cây lệch hoàn toàn."))
paras.append(para("    • Chỉ các BST tự cân bằng (AVL, Red-Black Tree) đảm bảo O(log n) mọi lúc."))
paras.append(para(""))

# ============================
# KẾT THÚC
# ============================
paras.append(heading("Dữ liệu cá nhân theo MSSV N24DCDT047", 1))
paras.append(para("Chạy: julia TH_2026_10_09_Sinh_vien.jl --data N24DCDT047"))
paras.append(para("      julia TH_2026_10_09_Sinh_vien.jl N24DCDT047"))
paras.append(para("Ngân sách B = 30  |  Số đơn n = 9"))
paras.append(para(""))
paras.append(table(
    table_row("ID", "Điểm giao", "Phục vụ", "Chi phí τ", "Lợi ích", "Tuyến đi", header=True),
    table_row("1", "5", "1", "5", "10", "1→3→5"),
    table_row("2", "6", "3", "9", "12", "1→3→5→6"),
    table_row("3", "3", "2", "4", "8", "1→3"),
    table_row("4", "6", "1", "7", "7", "1→3→5→6"),
    table_row("5", "5", "3", "7", "12", "1→3→5"),
    table_row("6", "4", "2", "6", "6", "1→3→4"),
    table_row("7", "2", "2", "4", "15", "1→2"),
    table_row("8", "6", "2", "8", "5", "1→3→5→6"),
    table_row("9", "3", "1", "3", "12", "1→3"),
))
paras.append(para(""))
paras.append(para("Kết quả DP cá nhân: ids=[9,7,3,1,5,4], cost=3+4+4+5+7+7=30, value=12+15+8+10+12+7=64"))
paras.append(para("DP: 64 | Vét cạn: 64 ✅ | Tham lam: 63 (minh chứng tham lam không luôn tối ưu)"))

# ============================
# Ghép thành body XML
# ============================
body_content = '\n'.join(paras)

DOCUMENT = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
            xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
            xmlns:o="urn:schemas-microsoft-com:office:office"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
            xmlns:v="urn:schemas-microsoft-com:vml"
            xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
            xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
            xmlns:w10="urn:schemas-microsoft-com:office:word"
            xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
            xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
            xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
            xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
            xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
            mc:Ignorable="w14 wp14">
<w:body>
{body_content}
<w:sectPr>
  <w:pgSz w:w="12240" w:h="15840"/>
  <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>'''

# ===========================================================================
# Ghi file docx (zip)
# ===========================================================================
with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', CONTENT_TYPES)
    z.writestr('_rels/.rels', RELS)
    z.writestr('word/_rels/document.xml.rels', WORD_RELS)
    z.writestr('word/styles.xml', STYLES)
    z.writestr('word/settings.xml', SETTINGS)
    z.writestr('word/document.xml', DOCUMENT.encode('utf-8'))

print(f"✅ Đã tạo file: {output_path}")
print(f"   Kích thước: {os.path.getsize(output_path):,} bytes")

