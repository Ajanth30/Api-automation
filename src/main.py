import yaml
from excel_postman_generator import generate_postman_collection_from_excel
from emailer import send_results_email
from auth_client import fetch_auth_token, AuthError


def main():
    """Main entry point: load config and generate Postman collections."""
    with open("services_config.yaml", 'r') as f:
        config = yaml.safe_load(f)

    # Excel-driven mode (no Swagger)
    excel_path = config.get("excel_path")
    collection_name = config.get("collection_name", "API Tests")
    gateway_base_url = config.get("gateway_base_url")
    if excel_path:
        auth_cfg = (config or {}).get("auth") or {}
        auth_headers = None
        auth_info = None
        if auth_cfg:
            auth_base_url = auth_cfg.get("base_url") or gateway_base_url
            try:
                token = fetch_auth_token(auth_base_url, auth_cfg)
                header_name = auth_cfg.get("header_name", "Authorization")
                prefix = auth_cfg.get("header_prefix", "Bearer ")
                if prefix is None:
                    header_value = token
                else:
                    header_value = f"{prefix}{token}" if prefix else token
                auth_headers = {header_name: header_value}
                if header_name.lower() == "authorization" and (prefix or "").strip().lower().startswith("bearer"):
                    bearer_token = header_value[len(prefix):] if prefix else token
                    auth_info = {
                        "type": "bearer",
                        "token": bearer_token or token,
                    }
                print("🔐 Auth token fetched and will be applied to all requests.")
            except AuthError as exc:
                print(f"❌ Authentication failed: {exc}")
                return

        print(f"\n📄 Excel mode enabled. Reading tests from: {excel_path}")
        collection_file, results_excel, failed_ids = generate_postman_collection_from_excel(
            excel_path,
            collection_name,
            base_url_override=gateway_base_url,
            auth_headers=auth_headers,
            auth_info=auth_info,
        )
        print(f"📦 Collection ready: {collection_file}")
        if results_excel:
            print(f"📘 Test results: {results_excel}")

        # Optional email notification
        email_cfg = (config or {}).get("email") or {}
        recipients = email_cfg.get("recipients") or []
        smtp_cfg = email_cfg.get("smtp") or {}
        subject = email_cfg.get("subject") or f"API test results: {collection_name}"

        if recipients:
            failed_list_text = "\n".join(f"- {fid}" for fid in failed_ids) if failed_ids else "- None"
            body = (
                f"Failed test case IDs ({len(failed_ids)}):\n{failed_list_text}\n"
            )
            send_results_email(
                recipients=recipients,
                subject=subject,
                body_text=body,
                attachments=[p for p in [collection_file, results_excel] if p],
                smtp=smtp_cfg,
                sender=email_cfg.get("from"),
            )
        return

if __name__ == "__main__":
    main()


