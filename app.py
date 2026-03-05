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

    # --- Data preview ---
    st.subheader("Data Preview")
    st.caption(f"Showing data from sheet: **{sheet_name}**")
    st.dataframe(df.head())

    # --- Plot type selection ---
    st.sidebar.header("Plot Type")
    plot_type = st.sidebar.radio(
        "Select the plot type",
        ["Scatter Plot", "Stacked Bar Chart", "Step Line Plot"],
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
        x_col = st.sidebar.selectbox("X-axis column", numeric_cols, key="scatter_x")
        y_col = st.sidebar.selectbox("Y-axis column", numeric_cols, key="scatter_y")
        x_label = st.sidebar.text_input("X-axis label", value=x_col, key="scatter_xlab")
        y_label = st.sidebar.text_input("Y-axis label", value=y_col, key="scatter_ylab")
        title = st.sidebar.text_input("Plot title", value=f"{y_col} vs {x_col}", key="scatter_title")

    elif plot_type == "Stacked Bar Chart":
        x_label = st.sidebar.text_input("X-axis label", value="Row Label", key="bar_xlab")
        y_label = st.sidebar.text_input("Y-axis label", value="Value", key="bar_ylab")
        title = st.sidebar.text_input("Plot title", value="Stacked Bar Chart", key="bar_title")

    else:  # Step line plot
        if len(numeric_cols) < 2:
            st.error("Step line plot requires at least TWO numeric columns.")
            st.stop()

        st.sidebar.subheader("Step Plot Options")
        sheets_to_plot = st.sidebar.multiselect(
            "Sheets to include (one step line per sheet)",
            options=sheet_names,
            default=sheet_names,
            key="step_sheets"
        )
        if len(sheets_to_plot) == 0:
            st.warning("Please select at least one worksheet.")
            st.stop()

        x_col = st.sidebar.selectbox("X-axis column", numeric_cols, key="step_x")
        y_col = st.sidebar.selectbox("Y-axis column", numeric_cols, key="step_y")
        x_label = st.sidebar.text_input("X-axis label", value=x_col, key="step_xlab")
        y_label = st.sidebar.text_input("Y-axis label", value=y_col, key="step_ylab")
        title = st.sidebar.text_input("Plot title", value=f"Step: {y_col} vs {x_col}", key="step_title")
        linewidth = st.sidebar.number_input("Line width", 0.5, 10.0, 2.0, 0.5, key="step_lw")

    # Figure size
    col1, col2 = st.sidebar.columns(2)
    with col1:
        fig_width = st.number_input("Figure width", 2.0, 20.0, 10.0)
    with col2:
        fig_height = st.number_input("Figure height", 2.0, 20.0, 6.0)

    # Font size
    st.sidebar.header("Font Settings")
    font_size = st.sidebar.number_input("Global font size", 6, 40, 12, 1)
    plt.rcParams["font.size"] = font_size

    # --- Axis Scale (shared by Scatter + Step)
    if plot_type in ["Scatter Plot", "Step Line Plot"]:
        st.sidebar.header("Axis Scale")
        x_scale = st.sidebar.selectbox("X-axis scale", ["Linear", "Log"], index=0)
        y_scale = st.sidebar.selectbox("Y-axis scale", ["Linear", "Log"], index=0)
    else:
        x_scale = "Linear"
        y_scale = "Linear"

    # ============================================================
    #                        SCATTER PLOT
    # ============================================================
    if plot_type == "Scatter Plot":

        st.sidebar.header("Fit Options")
        fit_type = st.sidebar.selectbox(
            "Fit type",
            ["None", "Linear", "Polynomial", "Exponential"]
        )

        poly_degree = None
        if fit_type == "Polynomial":
            poly_degree = st.sidebar.number_input(
                "Polynomial degree", 2, 10, 2, 1
            )

        x = df[x_col].values
        y = df[y_col].values

        fig, ax = plt.subplots(figsize=(fig_width, fig_height), dpi=300)
        ax.plot(x, y, "k.", alpha=0.8)

        mask = ~np.isnan(x) & ~np.isnan(y)
        x_clean = x[mask]
        y_clean = y[mask]

        # --- Fits ---
        if fit_type == "Linear" and len(x_clean) > 1:
            m, b = np.polyfit(x_clean, y_clean, 1)
            x_fit = np.linspace(x_clean.min(), x_clean.max(), 300)
            y_fit = m * x_fit + b
            ax.plot(x_fit, y_fit, "r", linewidth=2)
            ax.legend([f"Linear Fit: y={m:.3g}x+{b:.3g}"])

        elif fit_type == "Polynomial" and poly_degree and len(x_clean) > poly_degree:
            coeffs = np.polyfit(x_clean, y_clean, poly_degree)
            poly = np.poly1d(coeffs)
            x_fit = np.linspace(x_clean.min(), x_clean.max(), 300)
            y_fit = poly(x_fit)
            ax.plot(x_fit, y_fit, "r", linewidth=2)
            ax.legend([f"Poly deg {poly_degree}"])

        elif fit_type == "Exponential" and len(x_clean) > 1:
            pos_mask = y_clean > 0
            if pos_mask.sum() > 1:
                x_pos = x_clean[pos_mask]
                y_pos = y_clean[pos_mask]
                b, log_a = np.polyfit(x_pos, np.log(y_pos), 1)
                a = np.exp(log_a)
                x_fit = np.linspace(x_pos.min(), x_pos.max(), 300)
                ax.plot(x_fit, a * np.exp(b * x_fit), "r", linewidth=2)
                ax.legend([f"Exp Fit"])
            else:
                st.warning("Exponential fit requires positive y-values.")

        # --- Apply labels & scaling ---
        ax.set_xlabel(x_label)
        ax.set_ylabel(y_label)
        ax.set_title(title)
        ax.grid(True, linestyle="--", alpha=0.4)

        if x_scale == "Log":
            ax.set_xscale("log")
        if y_scale == "Log":
            ax.set_yscale("log")

        st.pyplot(fig)

        # Export
        st.subheader("Export Plot")

        def get_image_bytes(fmt="png"):
            buf = io.BytesIO()
            fig.savefig(buf, format=fmt, dpi=300, bbox_inches="tight")
            buf.seek(0)
            return buf

        col_png, col_eps = st.columns(2)
        with col_png:
            st.download_button("Download PNG", get_image_bytes("png"),
                               "plot.png", "image/png")
        with col_eps:
            st.download_button("Download EPS", get_image_bytes("eps"),
                               "plot.eps", "application/postscript")

    # ============================================================
    #                   STACKED BAR CHART
    # ============================================================
    elif plot_type == "Stacked Bar Chart":

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
            vals = df_bar[col].fillna(0).to_numpy()
            ax_bar.bar(indices, vals, bottom=cumulative, label=col)
            cumulative += vals

        ax_bar.set_xticks(indices)
        ax_bar.set_xticklabels(x_labels, rotation=45, ha="right")
        ax_bar.set_xlabel(x_label)
        ax_bar.set_ylabel(y_label)
        ax_bar.set_title(title)
        ax_bar.grid(True, linestyle="--", alpha=0.4)
        ax_bar.legend()

        st.pyplot(fig_bar)

        # Export
        st.subheader("Export Bar Chart")

        def get_bar_bytes(fmt="png"):
            buf = io.BytesIO()
            fig_bar.savefig(buf, format=fmt, dpi=300, bbox_inches="tight")
            buf.seek(0)
            return buf

        col_png_bar, col_eps_bar = st.columns(2)
        with col_png_bar:
            st.download_button("Download PNG", get_bar_bytes("png"),
                               "bar_chart.png", "image/png")
        with col_eps_bar:
            st.download_button("Download EPS", get_bar_bytes("eps"),
                               "bar_chart.eps", "application/postscript")

    # ============================================================
    #                      STEP LINE PLOT
    # ============================================================
    else:
        st.subheader("Step Line Plot (one line per worksheet)")

        fig_step, ax_step = plt.subplots(figsize=(fig_width, fig_height), dpi=300)

        plotted_any = False
        skipped = []

        for sheet in sheets_to_plot:
            try:
                df_s = pd.read_excel(xlsx, sheet_name=sheet)

                # Numeric conversion
                x_vals = pd.to_numeric(df_s[x_col], errors="coerce").to_numpy()
                y_vals = pd.to_numeric(df_s[y_col], errors="coerce").to_numpy()

                mask = ~np.isnan(x_vals) & ~np.isnan(y_vals)
                x_clean = x_vals[mask]
                y_clean = y_vals[mask]

                if len(x_clean) == 0:
                    skipped.append(f"{sheet} (no valid data)")
                    continue

                # Sort by X
                order = np.argsort(x_clean)
                x_sorted = x_clean[order]
                y_sorted = y_clean[order]

                ax_step.step(x_sorted, y_sorted, where="pre",
                             linewidth=linewidth, label=sheet)
                plotted_any = True

            except KeyError:
                skipped.append(f"{sheet} (missing '{x_col}' or '{y_col}')")
            except Exception as e:
                skipped.append(f"{sheet} (error: {e})")

        ax_step.set_xlabel(x_label)
        ax_step.set_ylabel(y_label)
        ax_step.set_title(title)
        ax_step.grid(True, linestyle="--", alpha=0.4)

        # Apply scaling
        if x_scale == "Log":
            ax_step.set_xscale("log")
        if y_scale == "Log":
            ax_step.set_yscale("log")

        if plotted_any and len(sheets_to_plot) > 1:
            ax_step.legend()

        if skipped:
            st.warning("Some sheets were skipped:\n- " + "\n- ".join(skipped))

        if plotted_any:
            st.pyplot(fig_step)
        else:
            st.error("No valid data to plot.")
            st.stop()

        # Export
        st.subheader("Export Step Plot")

        def get_step_bytes(fmt="png"):
            buf = io.BytesIO()
            fig_step.savefig(buf, format=fmt, dpi=300, bbox_inches="tight")
            buf.seek(0)
            return buf

        col_png_s, col_eps_s = st.columns(2)
        with col_png_s:
            st.download_button("Download PNG", get_step_bytes("png"),
                               "step_plot.png", "image/png")
        with col_eps_s:
            st.download_button("Download EPS", get_step_bytes("eps"),
                               "step_plot.eps", "application/postscript")

else:
    st.info("Upload an Excel file to begin.")