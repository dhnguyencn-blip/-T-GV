import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Phần mềm Theo dõi & Tổng hợp Điểm trừ Giáo viên",
    page_icon="📊",
    layout="wide",
)

# Khởi tạo dữ liệu mẫu hoặc đọc từ file Excel gốc nếu có
@st.cache_data
def load_initial_data():
  try:
    # Cố gắng đọc từ file Excel có sẵn
    excel_path = 'BANG_CHAM_CONG_THEO_THANG 9.xlsx'
    df_cc = pd.read_excel(excel_path, sheet_name='Chấm công tháng', skiprows=3)
    df_cc = df_cc.dropna(subset=['Họ và tên'])
    # Lọc các cột cần thiết
    cols = [
        'Họ và tên',
        'Chức vụ',
        'Vắng',
        'Không chào cờ',
        'Vắng giao ban',
        'Vắng lao động',
        'Vắng trực bán trú',
        'Đi chậm 5 phút',
        'Đồng phục không đúng quy định',
        'Hút thuốc không đúng nơi',
    ]
    df_cc = df_cc[cols]
    for c in cols[2:]:
      df_cc[c] = pd.to_numeric(df_cc[c], errors='fillna').fillna(0).astype(int)

    df_qd = pd.read_excel(excel_path, sheet_name='Quy định điểm trừ', skiprows=1)
    df_qd.columns = ['Nội dung điểm trừ', 'Mức điểm trừ', 'Ghi chú']
    return df_cc, df_qd
  except Exception:
    # Dữ liệu mặc định nếu không đọc được file
    data_teachers = [
        ['Nguyễn Đăng Huy', 'HP', 0, 0, 0, 0, 0, 0, 0, 0],
        ['Trần Trọng Nghĩa', 'GV', 0, 0, 0, 0, 0, 0, 0, 0],
        ['Thái văn Khiêm', 'GV', 1, 0, 0, 0, 0, 0, 0, 0],
        ['Dương Hồng Nguyên', 'GV', 6, 0, 0, 0, 0, 0, 0, 0],
        ['Ngô Xuân An', 'GV', 0, 0, 0, 0, 0, 0, 0, 0],
        ['Hoàng Mạnh Quân', 'GV', 0, 0, 0, 0, 0, 0, 0, 0],
        ['Hồ Viết Kiên', 'GV', 0, 0, 0, 0, 0, 0, 0, 0],
        ['Lô Thị Ngoan', 'TB', 0, 0, 0, 0, 0, 0, 0, 0],
    ]
    cols = [
        'Họ và tên',
        'Chức vụ',
        'Vắng',
        'Không chào cờ',
        'Vắng giao ban',
        'Vắng lao động',
        'Vắng trực bán trú',
        'Đi chậm 5 phút',
        'Đồng phục không đúng quy định',
        'Hút thuốc không đúng nơi',
    ]
    df_cc = pd.DataFrame(data_teachers, columns=cols)

    data_rules = [
        ['Vắng', 1, 'Trừ điểm nghỉ không phép/có phép theo quy chế'],
        ['Không chào cờ', 1, 'Vi phạm nội quy sinh hoạt đầu tuần'],
        ['Vắng giao ban', 1, 'Không tham gia họp giao ban định kỳ'],
        ['Vắng lao động', 1, 'Trốn buổi lao động tập thể'],
        ['Vắng trực bán trú', 1, 'Bỏ nhiệm vụ trực bán trú học sinh'],
        ['Đi chậm 5 phút', 1, 'Đi làm muộn so với quy định'],
        ['Đồng phục không đúng quy định', 1, 'Không mặc đúng trang phục giáo viên'],
        ['Hút thuốc không đúng nơi', 1, 'Hút thuốc trong khu vực cấm'],
    ]
    df_qd = pd.DataFrame(
        data_rules, columns=['Nội dung điểm trừ', 'Mức điểm trừ', 'Ghi chú']
    )
    return df_cc, df_qd


if 'df_teachers' not in st.session_state:
  st.session_state['df_teachers'], st.session_state['df_rules'] = (
      load_initial_data()
  )

st.title(
    '🏫 ỨNG DỤNG THEO DÕI & TỔNG HỢP ĐIỂM TRỪ GIÁO VIÊN HÀNG THÁNG'
)
st.markdown(
    '**Đơn vị:** Trường PTDTBT THCS YÊN HÒA | **Hệ thống Quản lý Thi đua Tự'
    ' động**'
)

# Sidebar menu
menu = st.sidebar.selectbox(
    '📋 Chọn chức năng',
    [
        '1. Bảng Chấm công & Nhập liệu Tháng',
        '2. Quản lý Danh mục & Mức Điểm trừ',
        '3. Tổng hợp & Báo cáo Thống kê',
        '4. Quản lý Danh sách Giáo viên',
    ],
)

# Lấy mapping mức điểm trừ từ danh mục quy định
rules_dict = dict(
    zip(
        st.session_state['df_rules']['Nội dung điểm trừ'],
        st.session_state['df_rules']['Mức điểm trừ'],
    )
)
violation_cols = [
    c for c in st.session_state['df_teachers'].columns if c not in ['Họ và tên', 'Chức vụ']
]

# Tính toán tự động tổng điểm trừ và điểm còn lại cho bảng giáo viên
def calculate_scores(df):
  df_calc = df.copy()
  total_penalties = 0
  total_deductions = []
  remaining_scores = []
  total_violations = []

  for idx, row in df_calc.iterrows():
    v_count = 0
    d_score = 0
    for col in violation_cols:
      count = int(row.get(col, 0))
      v_count += count
      penalty_per_unit = rules_dict.get(col, 1)
      d_score += count * penalty_per_unit

    total_violations.append(v_count)
    total_deductions.append(d_score)
    remaining_scores.append(max(0, 100 - d_score))

  df_calc['Tổng lượt vi phạm'] = total_violations
  df_calc['Tổng điểm trừ'] = total_deductions
  df_calc['Điểm còn lại'] = remaining_scores
  return df_calc

df_processed = calculate_scores(st.session_state['df_teachers'])

if menu == '1. Bảng Chấm công & Nhập liệu Tháng':
  st.subheader('📌 Bảng Chấm công và Ghi nhận Lỗi Vi phạm trong Tháng')
  st.info(
      '💡 Hướng dẫn: Bạn có thể chỉnh sửa trực tiếp số lần vi phạm của giáo'
      ' viên vào các cột lỗi bên dưới. Hệ thống sẽ tự động tính toán tổng lượt'
      ' vi phạm, tổng điểm trừ và điểm còn lại (Thang điểm chuẩn: 100).'
  )

  # Cho phép chỉnh sửa trực tiếp bảng dữ liệu
  edited_df = st.data_editor(
      st.session_state['df_teachers'],
      num_rows='dynamic',
      key='teacher_editor',
      use_container_width=True,
  )

  # Cập nhật lại session state khi chỉnh sửa
  if not edited_df.equals(st.session_state['df_teachers']):
    st.session_state['df_teachers'] = edited_df
    st.rerun()

  st.markdown('---')
  st.subheader('📈 Bảng Tổng hợp Kết quả Sau khi Tính điểm')
  st.dataframe(df_processed, use_container_width=True)

  # Xuất file Excel
  col1, col2 = st.columns([1, 4])
  with col1:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
      df_processed.to_excel(writer, index=False, sheet_name='TongHopDiemTru')
    excel_data = output.getvalue()

    st.download_button(
        label='📥 Tải xuống Báo cáo Excel',
        data=excel_data,
        file_name='Tong_Hop_Diem_Tru_Thang.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    )

elif menu == '2. Quản lý Danh mục & Mức Điểm trừ':
  st.subheader('⚙️ Thiết lập Quy định & Mức Điểm trừ cho Từng Lỗi')
  st.markdown(
      'Bạn có thể thay đổi mức điểm trừ cho mỗi lần vi phạm hoặc thêm/sửa các lỗi'
      ' vi phạm theo quy chế của nhà trường.'
  )

  edited_rules = st.data_editor(
      st.session_state['df_rules'],
      num_rows='dynamic',
      key='rules_editor',
      use_container_width=True,
  )

  if not edited_rules.equals(st.session_state['df_rules']):
    st.session_state['df_rules'] = edited_rules
    st.rerun()

elif menu == '3. Tổng hợp & Báo cáo Thống kê':
  st.subheader('📊 Thống kê & Phân tích Thi đua Giáo viên')

  col1, col2, col3 = st.columns(3)
  col1.metric('Tổng số Cán bộ, Giáo viên', len(df_processed))
  col2.metric('Tổng lượt vi phạm trong tháng', int(df_processed['Tổng lượt vi phạm'].sum()))
  col3.metric(
      'Điểm trung bình toàn trường',
      f"{df_processed['Điểm còn lại'].mean():.1f} / 100",
  )

  st.markdown('### 🏆 Biểu đồ Điểm số Thi đua của Giáo viên')
  chart_data = df_processed.set_index('Họ và tên')[['Điểm còn lại']]
  st.bar_chart(chart_data)

  st.markdown('### 🔍 Danh sách Cán bộ, Giáo viên có Vi phạm (Điểm < 100)')
  violation_list = df_processed[df_processed['Tổng điểm trừ'] > 0]
  if len(violation_list) > 0:
    st.dataframe(violation_list[['Họ và tên', 'Chức vụ', 'Tổng lượt vi phạm', 'Tổng điểm trừ', 'Điểm còn lại']], use_container_width=True)
  else:
    st.success('🎉 Tuyệt vời! Tháng này không có cán bộ, giáo viên nào vi phạm.')

elif menu == '4. Quản lý Danh sách Giáo viên':
  st.subheader('👥 Danh sách Cán bộ, Giáo viên, Nhân viên')
  st.markdown('Thêm mới, xóa hoặc chỉnh sửa thông tin nhân sự của trường.')
  
  st.dataframe(st.session_state['df_teachers'][['Họ và tên', 'Chức vụ']], use_container_width=True)