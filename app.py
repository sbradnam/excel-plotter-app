import io
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

st.set_page_config(page_title="Interactive Plotter", layout="wide")
st.title("📈 Interactive Excel Plotter")

# --- File upload ---
uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:

    # --- Sheet selection ---
    try:
        xlsx = pd.ExcelFile(uploaded_file)
        sheet_names = xlsx.sheet_names
        st.sidebar.header("Excel Sheet")

        if len(sheet_names) > 1:
            sheet_name = st.sidebar.selectbox("Select Sheet", sheet_names, index=0)
        else:
            sheet_name = sheet_names[0]

        df = pd.read_excel(xlsx, sheet_name=sheet_name)

    except Exception as e:
        st.error(f"Error reading file: {e}")
        st.stop()

    st.subheader("Data Preview")
    st.caption(f"Showing data from sheet: **{sheet_name}**")
    st.dataframe(df.head())

    # --- Plot type selection ---
    st.sidebar.header("Plot Type")
    plot_type = st.sidebar.radio(
        "Select the plot type",
        ["Scatter Plot", "Stacked Bar Chart"],
        index=0
    )

    # --- Column selection ---
    st.sidebar.header("Column Selection")
    all_columns = df.columns.tolist()
    selected_columns = st.sidebar.multiselect(
        "Select columns to use",
        all_columns,
        default=all_columns
    )

    if len(selected_columns) < 1:
        st.warning("Please select at least one column.")
        st.stop()

    numeric_cols = df[selected_columns].select_dtypes(include=["number"]).columns.tolist()
    if len(numeric_cols) == 0:
        st.error("No numeric columns available.")
        st.stop()

    # --- Shared settings ---
    st.sidebar.header("Plot Settings")

    if plot_type == "Scatter Plot":
        if len(numeric_cols) < 2:
            st.error("Scatter plot requires at least TWO numeric columns.")
            st.stop()
        x_col = st.sidebar.selectbox("X-axis column", numeric_cols)
        y_col = st.sidebar.selectbox("Y-axis column", numeric_cols)

        x_label = st.sidebar.text_input("X-axis label", value=x_col)
        y_label = st.sidebar.text_input("Y-axis label", value=y_col)
        title = st.sidebar.text_input("Plot title", value=f"{y_col} vs {x_col}")

    else:
        x_label = st.sidebar.text_input("X-axis label", value="Row Label")
        y_label = st.sidebar.text_input("Y-axis label", value="Value")
        title = st.sidebar.text_input("Plot title", value="Stacked Bar Chart")

    col1, col2 = st.sidebar.columns(2)
    with col1:
        fig_width = st.number_input("Figure width", min_value=2.0, max_value=20.0, value=10.0)
    with col2:
        fig_height = st.number_input("Figure height", min_value=2.0, max_value=20.0, value=6.0)

    st.sidebar.header("Font Settings")
    font_size = st.sidebar.number_input(
        "Global font size", min_value=6, max_value=40, value=12, step=1
    )
    plt.rcParams["font.size"] = font_size

    # --- SCATTER PLOT MODE ---
    if plot_type == "Scatter Plot":

        st.sidebar.header("Fit Options")
        fit_type = st.sidebar.selectbox(
            "Fit type",
            ["None", "Linear", "Polynomial", "Exponential"]
        )

        poly_degree = None
        if fit_type == "Polynomial":
            poly_degree = st.sidebar.number_input(
                "Polynomial degree", min_value=2, max_value=10, value=2, step=1
            )

        x = df[x_col].values
        y = df[y_col].values

        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=300)
        ax.plot(x, y, "k.", alpha=0.8)

        mask = ~np.isnan(x) & ~np.isnan(y)
        x_clean = x[mask]
        y_clean = y[mask]

        if fit_type == "Linear" and len(x_clean) > 1:
            m, b = np.polyfit(x_clean, y_clean, 1)
            x_fit = np.linspace(x_clean.min(), x_clean.max(), 300)
            y_fit = m * x_fit + b
            ax.plot(x_fit, y_fit, color="red", linewidth=2,
                    label=f"Linear: y={m:.3g}x+{b:.3g}")

        elif fit_type == "Polynomial" and poly_degree is not None and len(x_clean) > poly_degree:
            coeffs = np.polyfit(x_clean, y_clean, poly_degree)
            poly = np.poly1d(coeffs)
            x_fit = np.linspace(x_clean.min(), x_clean.max(), 300)
            y_fit = poly(x_fit)

            terms = []
            deg = poly_degree
            for c in coeffs:
                if abs(c) < 1e-12:
                    deg -= 1
                    continue
                if deg == 0:
                    terms.append(f"{c:.3g}")
                elif deg == 1:
                    terms.append(f"{c:.3g}·x")
                else:
                    terms.append(f"{c:.3g}·x^{deg}")
                deg -= 1

            poly_str = " + ".join(terms).replace("+ -", "- ")
            ax.plot(x_fit, y_fit, color="red", linewidth=2,
                    label=f"Poly deg {poly_degree}: y={poly_str}")

        elif fit_type == "Exponential" and len(x_clean) > 1:
            positive_mask = y_clean > 0
            if positive_mask.sum() > 1:
                x_pos = x_clean[positive_mask]
                y_pos = y_clean[positive_mask]
                log_y = np.log(y_pos)
                b, log_a = np.polyfit(x_pos, log_y, 1)
                a = np.exp(log_a)
                x_fit = np.linspace(x_pos.min(), x_pos.max(), 300)
                y_fit = a * np.exp(b * x_fit)
                ax.plot(x_fit, y_fit, color="red", linewidth=2,
                        label=f"Exp: y={a:.3g}·e^{b:.3g}x")
            else:
                st.warning("Exponential fit requires positive y-values.")

        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        ax.set_title(title)
        ax.grid(True, linestyle="--", alpha=0.4)
        if fit_type != "None":
            ax.legend()

        st.pyplot(fig)

        # Export scatter
        st.subheader("Export Plot")

        def get_image_bytes(fmt="png"):
            buf = io.BytesIO()
            fig.savefig(buf, format=fmt, dpi=300, bbox_inches="tight")
            buf.seek(0)
            return buf

        col_png, col_eps = st.columns(2)
        with col_png:
            st.download_button(
                "Download PNG (300 dpi)",
                data=get_image_bytes("png"),
                file_name="plot.png",
                mime="image/png"
            )
        with col_eps:
            st.download_button(
                "Download EPS (300 dpi)",
                data=get_image_bytes("eps"),
                file_name="plot.eps",
                mime="application/postscript"
            )

    # --- STACKED BAR CHART MODE ---
    else:
        st.subheader("Stacked Bar Chart (Each Row)")

        if len(numeric_cols) < 2:
            st.error("Stacked bar chart requires at least TWO numeric columns.")
            st.stop()

        df_bar = df[numeric_cols]

        label_col = selected_columns[0]
        x_labels = df[label_col].astype(str)

        indices = np.arange(len(df_bar))

        fig_bar, ax_bar = plt.subplots(figsize=(fig_width, fig_height), dpi=300)
        cumulative = np.zeros(len(df_bar))

        for col in numeric_cols:
            values = df_bar[col].fillna(0).values
            ax_bar.bar(indices, values, bottom=cumulative, label=col)
            cumulative += values

        ax_bar.set_xticks(indices)
        ax_bar.set_xticklabels(x_labels, rotation=45, ha="right")

        ax_bar.set_xlabel(x_label)
        ax_bar.set_ylabel(y_label)
        ax_bar.set_title(title)
        ax_bar.legend()
        ax_bar.grid(True, linestyle="--", alpha=0.4)

        st.pyplot(fig_bar)

        # Export bar chart
        st.subheader("Export Bar Chart")

        def get_bar_image_bytes(fmt="png"):
            buf = io.BytesIO()
            fig_bar.savefig(buf, format=fmt, dpi=300, bbox_inches="tight")
            buf.seek(0)
            return buf

        col_png_bar, col_eps_bar = st.columns(2)
        with col_png_bar:
            st.download_button(
                "Download PNG (300 dpi)",
                data=get_bar_image_bytes("png"),
                file_name="bar_chart.png",
                mime="image/png"
            )

        with col_eps_bar:
            st.download_button(
                "Download EPS (300 dpi)",
                data=get_bar_image_bytes("eps"),
                file_name="bar_chart.eps",
                mime="application/postscript"
            )

else:
    st.info("Upload an Excel file to begin.")