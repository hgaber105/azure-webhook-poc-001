import os
from flask import Flask, jsonify, request

app = Flask(__name__)


@app.route("/", methods=["POST"])
def webhook_receiver():
  """Endpoint to receive incoming webhook events."""
  try:
    payload = request.get_json(silent=True)

    if not payload:
      return (
          jsonify({"status": "error", "message": "Invalid or missing JSON"}),
          400,
      )

    # Process or log the webhook payload
    print(f"Received webhook payload: {payload}")

    # Respond to the webhook sender acknowledging receipt
    return jsonify({"status": "success", "message": "Webhook received. ver 001"}), 200

  except Exception as e:
    print(f"Error processing webhook: {str(e)}")
    return jsonify({"status": "error", "message": str(e)}), 500


if __name__ == "__main__":
  port = int(os.environ.get("PORT", 8080))
  app.run(host="0.0.0.0", port=port)