@app.route('/download-zips')
def download_zips():
    acc_api = ACCAPI()
    result = acc_api.download_project_zips("Information Systems Workspace")

    if "error" in result:
        return jsonify(result), result["status_code"]

    zip_files = result["files"]
    if not zip_files:
        return jsonify({"error": "No ZIP files found."}), 404

    # Create a temporary ZIP archive
    temp_zip_path = tempfile.NamedTemporaryFile(suffix=".zip", delete=False).name

    with zipfile.ZipFile(temp_zip_path, "w", zipfile.ZIP_DEFLATED) as temp_zip:
        for file_path in zip_files:
            zip_filename = os.path.basename(file_path)
            temp_zip.write(file_path, zip_filename)

    # Send the archive to the user
    response = send_file(temp_zip_path, as_attachment=True, download_name="all_zips.zip")
    response.call_on_close(lambda: os.remove(temp_zip_path))  # Ensure cleanup after response is sent

    return response


@app.route('/get-zips', methods=['POST', 'GET'])
def get_zips():
    acc_api = ACCAPI()
    project = None
    if request.method == 'POST':
        data = request.get_json(silent=True)  # Avoids error if JSON is missing
        if not data:
            return jsonify({"error": "Invalid or missing JSON", "status_code": 400}), 400
        url = data.get('url')

    elif request.method == 'GET':
        url = request.args.get('url')  # Get URL from query parameters

    if url:
        print(f"Processing request for URL: {url}")

        # Extract Project ID using regex
        pattern = r"projects/([a-f0-9-]{36})"
        match = re.search(pattern, url)
        if not match:
            return jsonify({"error": "Project ID not found in the URL", "status_code": 400}), 400
        project_id = match.group(1)
        project = acc_api.call_api(f"construction/admin/v1/projects/{project_id}")



    if project is None:
        result = acc_api.get_project_files(file_types=["zip", "rar", "7z"])
    else:
        result = acc_api.get_project_files(project["name"], file_types=["zip", "rar", "7z"])
        print(project["name"])

    return jsonify(result)