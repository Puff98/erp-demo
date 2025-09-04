import os
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, render_template_string, request, redirect, url_for, send_file
from flask_sqlalchemy import SQLAlchemy

# ------------------- CONFIG -------------------
# Change this path to the folder where you want monthly Excel files stored.
EXPORT_DIR = "./exports"   # <-- change to your desired folder (absolute path recommended)
Path(EXPORT_DIR).mkdir(parents=True, exist_ok=True)

# Flask + SQLite DB
app = Flask(__name__)
app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///erp_demo.db"
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False
db = SQLAlchemy(app)


# ------------------- MODELS -------------------
class Customer(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    gst_no = db.Column(db.String(50))
    address = db.Column(db.String(400))
    mobile = db.Column(db.String(50))
    email = db.Column(db.String(200))


class Item(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    hsn_code = db.Column(db.String(50))
    material = db.Column(db.String(200))  # optional


class Inward(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    entry_date = db.Column(db.Date, default=datetime.utcnow)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("item.id"), nullable=False)
    dc_no_cust = db.Column(db.String(200))   # DC number from customer
    qty = db.Column(db.Float, default=0.0)
    rate = db.Column(db.Float, default=0.0)
    amt = db.Column(db.Float, default=0.0)


class Outward(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    entry_date = db.Column(db.Date, default=datetime.utcnow)
    customer_id = db.Column(db.Integer, db.ForeignKey("customer.id"), nullable=False)
    item_id = db.Column(db.Integer, db.ForeignKey("item.id"), nullable=False)
    dc_no_cust = db.Column(db.String(200))   # DC number from customer (to match inward)
    dc_unique_no_noncust = db.Column(db.String(200))  # your unique DC reference
    qty = db.Column(db.Float, default=0.0)


with app.app_context():
    db.create_all()


# ------------------- UTIL: Excel append/ensure -------------------
def month_filename_for(date_obj: datetime):
    return os.path.join(EXPORT_DIR, f"{date_obj.year}-{date_obj.month:02d}.xlsx")


def append_row_to_sheet(month_file: str, sheet_name: str, row_dict: dict):
    """
    Append row_dict as a new row to specified sheet in Excel monthly file.
    If file or sheet does not exist, create them.
    """
    if os.path.exists(month_file):
        # load existing sheet into DataFrame (if sheet exists)
        try:
            existing = pd.read_excel(month_file, sheet_name=sheet_name, engine="openpyxl")
            df = pd.concat([existing, pd.DataFrame([row_dict])], ignore_index=True)
        except (ValueError, KeyError):
            # sheet not found -> create new
            df = pd.DataFrame([row_dict])
    else:
        df = pd.DataFrame([row_dict])

    # Write all sheets back — we will preserve other sheets by reading them first
    # Read existing workbook sheets
    writer = pd.ExcelWriter(month_file, engine="openpyxl", mode="w")
    # If there are other sheets (like the other sheet), write them back to preserve
    other_sheet = "Inward" if sheet_name == "Outward" else "Outward"
    if os.path.exists(month_file):
        try:
            other_df = pd.read_excel(month_file, sheet_name=other_sheet, engine="openpyxl")
            other_df.to_excel(writer, sheet_name=other_sheet, index=False)
        except Exception:
            pass

    df.to_excel(writer, sheet_name=sheet_name, index=False)
    writer.close()


# ------------------- TEMPLATES -------------------
BASE_TEMPLATE = """
<!doctype html>
<html>
<head>
  <title>ERP Demo - Python</title>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <style> .small-col { max-width:200px; } </style>
</head>
<body class="container py-4">
  <h2 class="mb-3">ERP Demo</h2>

  <ul class="nav nav-tabs mb-3">
    <li class="nav-item"><a class="nav-link {% if tab=='master' %}active{% endif %}" href="{{ url_for('master') }}">Master (Customer/Item)</a></li>
    <li class="nav-item"><a class="nav-link {% if tab=='inward' %}active{% endif %}" href="{{ url_for('inward') }}">Inward</a></li>
    <li class="nav-item"><a class="nav-link {% if tab=='outward' %}active{% endif %}" href="{{ url_for('outward') }}">Outward</a></li>
    <li class="nav-item"><a class="nav-link {% if tab=='overall' %}active{% endif %}" href="{{ url_for('overall') }}">Overall</a></li>
    <li class="nav-item"><a class="nav-link {% if tab=='settings' %}active{% endif %}" href="{{ url_for('settings') }}">Settings</a></li>
  </ul>

  <div>
    {{ content|safe }}
  </div>
</body>
</html>
"""


# ------------------- ROUTES -------------------
@app.route("/")
def home():
    return redirect(url_for("master"))


# MASTER: create and list customers/items with delete
@app.route("/master", methods=["GET", "POST"])
def master():
    if request.method == "POST":
        mode = request.form.get("mode")
        if mode == "add_customer":
            c = Customer(
                name=request.form["name"],
                gst_no=request.form.get("gst_no"),
                address=request.form.get("address"),
                mobile=request.form.get("mobile"),
                email=request.form.get("email")
            )
            db.session.add(c)
            db.session.commit()
        elif mode == "add_item":
            it = Item(
                name=request.form["item_name"],
                hsn_code=request.form.get("hsn"),
                material=request.form.get("material")
            )
            db.session.add(it)
            db.session.commit()
        return redirect(url_for("master"))

    # delete actions handled by query params
    del_type = request.args.get("del")
    if del_type == "cust":
        cid = request.args.get("id")
        if cid:
            Customer.query.filter_by(id=int(cid)).delete()
            db.session.commit()
            return redirect(url_for("master"))
    if del_type == "item":
        iid = request.args.get("id")
        if iid:
            Item.query.filter_by(id=int(iid)).delete()
            db.session.commit()
            return redirect(url_for("master"))

    customers = Customer.query.order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()

    content = render_template_string(
        """
<div class="row">
  <div class="col-md-6">
    <h5>Add Customer</h5>
    <form method="post">
      <input type="hidden" name="mode" value="add_customer">
      <input name="name" class="form-control mb-2" placeholder="Customer name" required>
      <input name="gst_no" class="form-control mb-2" placeholder="GST No">
      <input name="address" class="form-control mb-2" placeholder="Address">
      <input name="mobile" class="form-control mb-2" placeholder="Mobile">
      <input name="email" class="form-control mb-2" placeholder="Email">
      <button class="btn btn-primary">Add Customer</button>
    </form>

    <hr>
    <h6>Customers</h6>
    <table class="table table-sm">
      <thead><tr><th>Name</th><th>GST</th><th>Mobile</th><th>Action</th></tr></thead>
      <tbody>
        {% for c in customers %}
          <tr>
            <td>{{c.name}}</td><td>{{c.gst_no}}</td><td>{{c.mobile}}</td>
            <td><a class="btn btn-sm btn-danger" href="{{ url_for('master') }}?del=cust&id={{c.id}}">Delete</a></td>
          </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>

  <div class="col-md-6">
    <h5>Add Item</h5>
    <form method="post">
      <input type="hidden" name="mode" value="add_item">
      <input name="item_name" class="form-control mb-2" placeholder="Item name" required>
      <input name="hsn" class="form-control mb-2" placeholder="HSN Code">
      <input name="material" class="form-control mb-2" placeholder="Material (optional)">
      <button class="btn btn-primary">Add Item</button>
    </form>

    <hr>
    <h6>Items</h6>
    <table class="table table-sm">
      <thead><tr><th>Name</th><th>HSN</th><th>Material</th><th>Action</th></tr></thead>
      <tbody>
        {% for it in items %}
          <tr>
            <td>{{it.name}}</td><td>{{it.hsn_code}}</td><td>{{it.material}}</td>
            <td><a class="btn btn-sm btn-danger" href="{{ url_for('master') }}?del=item&id={{it.id}}">Delete</a></td>
          </tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
""",
        customers=customers, items=items
    )
    return render_template_string(BASE_TEMPLATE, content=content, tab="master")


# INWARD: form + save to DB + write to monthly Excel
@app.route("/inward", methods=["GET", "POST"])
def inward():
    customers = Customer.query.order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()

    if request.method == "POST":
        date_str = request.form.get("entry_date") or datetime.utcnow().strftime("%Y-%m-%d")
        entry_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        customer_id = int(request.form["customer_id"])
        item_id = int(request.form["item_id"])
        dc_no_cust = request.form.get("dc_no_cust")
        qty = float(request.form.get("qty") or 0)
        rate = float(request.form.get("rate") or 0)
        amt = qty * rate

        record = Inward(
            entry_date=entry_date,
            customer_id=customer_id,
            item_id=item_id,
            dc_no_cust=dc_no_cust,
            qty=qty,
            rate=rate,
            amt=amt
        )
        db.session.add(record)
        db.session.commit()

        # append to monthly excel
        month_file = month_filename_for(entry_date)
        # fetch details for human readable values
        cust = Customer.query.get(customer_id)
        it = Item.query.get(item_id)
        row = {
            "entry_date": entry_date.strftime("%Y-%m-%d"),
            "customer_name": cust.name if cust else "",
            "customer_gst": cust.gst_no if cust else "",
            "item_name": it.name if it else "",
            "hsn_code": it.hsn_code if it else "",
            "dc_no_cust": dc_no_cust,
            "qty": qty,
            "rate": rate,
            "amt": amt
        }
        append_row_to_sheet(month_file, "Inward", row)
        return redirect(url_for("inward"))

    content = render_template_string("""
<div class="row">
  <div class="col-md-8">
    <h5>Inward Entry</h5>
    <form method="post">
      <input type="date" name="entry_date" class="form-control mb-2" value="{{today}}">
      <select name="customer_id" class="form-control mb-2" required>
        <option value="">-- Select Customer --</option>
        {% for c in customers %}<option value="{{c.id}}">{{c.name}}</option>{% endfor %}
      </select>
      <select name="item_id" class="form-control mb-2" required>
        <option value="">-- Select Item --</option>
        {% for it in items %}<option value="{{it.id}}">{{it.name}} (HSN: {{it.hsn_code}})</option>{% endfor %}
      </select>
      <input name="dc_no_cust" class="form-control mb-2" placeholder="DC No (Customer)">
      <input name="qty" type="number" step="any" class="form-control mb-2" placeholder="Qty / Nos" required>
      <input name="rate" type="number" step="any" class="form-control mb-2" placeholder="Rate">
      <button class="btn btn-primary">Save Inward</button>
    </form>
  </div>

  <div class="col-md-4">
    <h6>Recent Inward (last 10)</h6>
    <table class="table table-sm">
      <thead><tr><th>Date</th><th>Cust</th><th>Item</th><th>Qty</th></tr></thead>
      <tbody>
        {% for r in recent %}
          <tr><td>{{r.entry_date}}</td><td>{{r.customer_name}}</td><td>{{r.item_name}}</td><td>{{r.qty}}</td></tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
""", customers=customers, items=items, today=datetime.utcnow().strftime("%Y-%m-%d"),
                                         recent=[
                                             {
                                                 "entry_date": inv.entry_date.strftime("%Y-%m-%d"),
                                                 "customer_name": Customer.query.get(inv.customer_id).name if Customer.query.get(inv.customer_id) else "",
                                                 "item_name": Item.query.get(inv.item_id).name if Item.query.get(inv.item_id) else "",
                                                 "qty": inv.qty
                                             }
                                             for inv in Inward.query.order_by(Inward.id.desc()).limit(10).all()
                                         ])
    return render_template_string(BASE_TEMPLATE, content=content, tab="inward")


# OUTWARD: form + save + excel
@app.route("/outward", methods=["GET", "POST"])
def outward():
    customers = Customer.query.order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()

    if request.method == "POST":
        date_str = request.form.get("entry_date") or datetime.utcnow().strftime("%Y-%m-%d")
        entry_date = datetime.strptime(date_str, "%Y-%m-%d").date()
        customer_id = int(request.form["customer_id"])
        item_id = int(request.form["item_id"])
        dc_no_cust = request.form.get("dc_no_cust")
        dc_unique = request.form.get("dc_unique")
        qty = float(request.form.get("qty") or 0)

        record = Outward(
            entry_date=entry_date,
            customer_id=customer_id,
            item_id=item_id,
            dc_no_cust=dc_no_cust,
            dc_unique_no_noncust=dc_unique,
            qty=qty
        )
        db.session.add(record)
        db.session.commit()

        # append to monthly excel
        month_file = month_filename_for(entry_date)
        cust = Customer.query.get(customer_id)
        it = Item.query.get(item_id)
        row = {
            "entry_date": entry_date.strftime("%Y-%m-%d"),
            "customer_name": cust.name if cust else "",
            "customer_gst": cust.gst_no if cust else "",
            "item_name": it.name if it else "",
            "hsn_code": it.hsn_code if it else "",
            "dc_no_cust": dc_no_cust,
            "dc_unique_no_noncust": dc_unique,
            "qty": qty
        }
        append_row_to_sheet(month_file, "Outward", row)
        return redirect(url_for("outward"))

    content = render_template_string("""
<div class="row">
  <div class="col-md-8">
    <h5>Outward Entry</h5>
    <form method="post">
      <input type="date" name="entry_date" class="form-control mb-2" value="{{today}}">
      <select name="customer_id" class="form-control mb-2" required>
        <option value="">-- Select Customer --</option>
        {% for c in customers %}<option value="{{c.id}}">{{c.name}}</option>{% endfor %}
      </select>
      <select name="item_id" class="form-control mb-2" required>
        <option value="">-- Select Item --</option>
        {% for it in items %}<option value="{{it.id}}">{{it.name}} (HSN: {{it.hsn_code}})</option>{% endfor %}
      </select>
      <input name="dc_no_cust" class="form-control mb-2" placeholder="DC No (Customer)">
      <input name="dc_unique" class="form-control mb-2" placeholder="DC Unique No (Non-Cust)">
      <input name="qty" type="number" step="any" class="form-control mb-2" placeholder="Qty / Nos" required>
      <button class="btn btn-primary">Save Outward</button>
    </form>
  </div>

  <div class="col-md-4">
    <h6>Recent Outward (last 10)</h6>
    <table class="table table-sm">
      <thead><tr><th>Date</th><th>Cust</th><th>Item</th><th>Qty</th></tr></thead>
      <tbody>
        {% for r in recent %}
          <tr><td>{{r.entry_date}}</td><td>{{r.customer_name}}</td><td>{{r.item_name}}</td><td>{{r.qty}}</td></tr>
        {% endfor %}
      </tbody>
    </table>
  </div>
</div>
""", customers=customers, items=items, today=datetime.utcnow().strftime("%Y-%m-%d"),
                                         recent=[
                                             {
                                                 "entry_date": out.entry_date.strftime("%Y-%m-%d"),
                                                 "customer_name": Customer.query.get(out.customer_id).name if Customer.query.get(out.customer_id) else "",
                                                 "item_name": Item.query.get(out.item_id).name if Item.query.get(out.item_id) else "",
                                                 "qty": out.qty
                                             }
                                             for out in Outward.query.order_by(Outward.id.desc()).limit(10).all()
                                         ])
    return render_template_string(BASE_TEMPLATE, content=content, tab="outward")


# OVERALL: show a combined view, filters for customer / item, compute dispatched & pending by dc_no_cust
@app.route("/overall")
def overall():
    cust_filter = request.args.get("customer_id", type=int)
    item_filter = request.args.get("item_id", type=int)

    # load filtered records
    inward_q = Inward.query
    outward_q = Outward.query
    if cust_filter:
        inward_q = inward_q.filter_by(customer_id=cust_filter)
        outward_q = outward_q.filter_by(customer_id=cust_filter)
    if item_filter:
        inward_q = inward_q.filter_by(item_id=item_filter)
        outward_q = outward_q.filter_by(item_id=item_filter)

    inwards = inward_q.order_by(Inward.entry_date.desc()).all()
    outwards = outward_q.order_by(Outward.entry_date.desc()).all()

    # Build DC mapping: for each dc_no_cust -> total inward qty, total outward qty
    dc_map = {}
    for inv in Inward.query.all():
        key = (inv.dc_no_cust or "").strip()
        if not key:
            continue
        m = dc_map.setdefault(key, {"inward": 0.0, "outward": 0.0})
        m["inward"] += float(inv.qty or 0.0)
    for out in Outward.query.all():
        key = (out.dc_no_cust or "").strip()
        if not key:
            continue
        m = dc_map.setdefault(key, {"inward": 0.0, "outward": 0.0})
        m["outward"] += float(out.qty or 0.0)

    # For the UI we will show rows for inwards (primary) and compute dispatched/pending from dc_map
    rows = []
    for inv in inwards:
        cust = Customer.query.get(inv.customer_id)
        it = Item.query.get(inv.item_id)
        key = (inv.dc_no_cust or "").strip()
        dispatched = dc_map.get(key, {}).get("outward", 0.0)
        inward_total = dc_map.get(key, {}).get("inward", inv.qty or 0.0)
        pending = inward_total - dispatched
        rows.append({
            "date": inv.entry_date.strftime("%Y-%m-%d"),
            "desc": "",  # you can use fields to fill description if needed
            "customer": cust.name if cust else "",
            "item": it.name if it else "",
            "hsn": it.hsn_code if it else "",
            "dc_no_cust": key,
            "qty_dispatch": dispatched,
            "pending_qty": pending
        })

    customers = Customer.query.order_by(Customer.name).all()
    items = Item.query.order_by(Item.name).all()

    content = render_template_string("""
<div class="mb-3">
  <form class="row g-2">
    <div class="col-auto">
      <select class="form-select" name="customer_id" onchange="this.form.submit()">
        <option value="">All Customers</option>
        {% for c in customers %}
          <option value="{{c.id}}" {% if request.args.get('customer_id', type=int) == c.id %}selected{% endif %}>{{c.name}}</option>
        {% endfor %}
      </select>
    </div>
    <div class="col-auto">
      <select class="form-select" name="item_id" onchange="this.form.submit()">
        <option value="">All Items</option>
        {% for it in items %}
          <option value="{{it.id}}" {% if request.args.get('item_id', type=int) == it.id %}selected{% endif %}>{{it.name}}</option>
        {% endfor %}
      </select>
    </div>
    <div class="col-auto">
      <a class="btn btn-secondary" href="{{ url_for('download_current_month') }}">Download current month Excel</a>
    </div>
  </form>
</div>

<table class="table table-sm table-bordered">
  <thead>
    <tr>
      <th>Date</th><th>Desc</th><th>Customer</th><th>Item</th><th>HSN</th><th>DC No (Cust)</th><th>Qty Dispatched</th><th>Pending Qty</th>
    </tr>
  </thead>
  <tbody>
    {% for r in rows %}
      <tr>
        <td>{{r.date}}</td><td>{{r.desc}}</td><td>{{r.customer}}</td><td>{{r.item}}</td><td>{{r.hsn}}</td>
        <td>{{r.dc_no_cust}}</td><td>{{r.qty_dispatch}}</td><td>{{r.pending_qty}}</td>
      </tr>
    {% endfor %}
  </tbody>
</table>
""", customers=customers, items=items, rows=rows)
    return render_template_string(BASE_TEMPLATE, content=content, tab="overall")


# Download current month file (if exists)
@app.route("/download-current")
def download_current_month():
    now = datetime.utcnow()
    f = month_filename_for(now)
    if os.path.exists(f):
        return send_file(f, as_attachment=True)
    else:
        return "Monthly Excel not found for current month.", 404


# SETTINGS: show current export dir and let user change (optional)
@app.route("/settings", methods=["GET", "POST"])
def settings():
    global EXPORT_DIR
    message = ""
    if request.method == "POST":
        new_path = request.form.get("export_dir", "").strip()
        if new_path:
            EXPORT_DIR = new_path
            Path(EXPORT_DIR).mkdir(parents=True, exist_ok=True)
            message = f"Export folder set to: {EXPORT_DIR}"
    content = render_template_string("""
<div>
  <h5>Settings</h5>
  <form method="post">
    <label>Excel Export Folder (absolute or relative)</label>
    <input name="export_dir" class="form-control mb-2" value="{{current}}">
    <button class="btn btn-primary">Save</button>
  </form>
  <div class="mt-3">
    <strong>Current Export Folder:</strong> {{current}}<br>
    <small class="text-muted">Monthly files: YYYY-MM.xlsx. Sheets: Inward, Outward</small>
  </div>
  {% if message %}
    <div class="alert alert-success mt-2">{{message}}</div>
  {% endif %}
</div>
""", current=EXPORT_DIR, message=message)
    return render_template_string(BASE_TEMPLATE, content=content, tab="settings")

# EXPORT: list all Excel files and allow download
@app.route("/export")
def export_files():
    files = []
    for f in sorted(os.listdir(EXPORT_DIR)):
        if f.endswith(".xlsx"):
            files.append(f)
    content = render_template_string("""
<h5>Exported Excel Files</h5>
{% if files %}
  <ul class="list-group">
    {% for f in files %}
      <li class="list-group-item d-flex justify-content-between align-items-center">
        {{f}}
        <a href="{{ url_for('download_file', filename=f) }}" class="btn btn-sm btn-primary">Download</a>
      </li>
    {% endfor %}
  </ul>
{% else %}
  <div class="alert alert-warning">No export files found yet.</div>
{% endif %}
""", files=files)
    return render_template_string(BASE_TEMPLATE, content=content, tab="export")


@app.route("/download/<path:filename>")
def download_file(filename):
    filepath = os.path.join(EXPORT_DIR, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True)
    return "File not found", 404

# ------------------- RUN -------------------
if __name__ == "__main__":
    import os as _os

    port = int(_os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
