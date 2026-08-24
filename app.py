import base64
import os

from flask import Flask, request, jsonify
from flask_cors import CORS

import uno_engine

app = Flask(__name__)
CORS(app)  # 給前端 (GitHub Pages) 跨網域呼叫用，個人使用暫不限制來源


@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


@app.route("/api/<product_code>/calculate", methods=["POST"])
def calculate(product_code):
    product_code = product_code.upper()
    if product_code not in uno_engine.PRODUCT_TEMPLATE:
        return jsonify({"error": f"未支援的商品代碼: {product_code}"}), 400

    data = request.get_json(force=True, silent=True) or {}

    required = ["gender", "birth_year", "birth_month", "birth_day", "payment_term"]
    missing = [f for f in required if f not in data or data[f] in (None, "")]
    if missing:
        return jsonify({"error": f"缺少必填欄位: {', '.join(missing)}"}), 400

    want_pdf = bool(data.get("want_pdf", False))

    try:
        result, pdf_path = uno_engine.calculate(product_code, data, want_pdf=want_pdf)
    except uno_engine.CalcError as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        return jsonify({"error": f"試算時發生錯誤: {e}"}), 500

    if pdf_path:
        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()
        os.remove(pdf_path)
        result["pdf_base64"] = base64.b64encode(pdf_bytes).decode("ascii")

    return jsonify(result)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
