from config.zotero_config import load_zotero_config

if __name__ == "__main__":
    cfg = load_zotero_config(
        allow_prompt=True,
        parent=None   # erzeugt eigenes Tk-Root
    )
    print("GELADENE KONFIGURATION:")
    print(cfg)
