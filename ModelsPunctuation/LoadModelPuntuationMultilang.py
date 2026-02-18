from huggingface_hub import snapshot_download

snapshot_download(
    "oliverguhr/fullstop-punctuation-multilang-large",
    repo_type="model",
    local_dir="./fullstop-punctuation-multilang-large"
)
