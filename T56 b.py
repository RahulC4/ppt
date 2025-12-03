# ----------------------------
# SOURCE PPT LIST + DELETE (UI SUPPORT)
# ----------------------------

def list_source_ppt_blobs():
    """
    List all source PPT files stored in the dataset container (ppt-dataset).
    Used by UI to show available templates.
    """
    try:
        container_client = _get_container_client(SOURCE_CONTAINER)
        return [b.name for b in container_client.list_blobs() if b.name.lower().endswith(".pptx")]
    except Exception as e:
        logger.warning(f"Failed to list source PPTs: {e}")
        return []


def delete_source_ppt_from_blob(blob_name: str):
    """
    Delete a source PPT from the dataset container (ppt-dataset).
    This is triggered when user removes a template from the UI.
    """
    try:
        container_client = _get_container_client(SOURCE_CONTAINER)
        container_client.delete_blob(blob_name)
        logger.info(f"Deleted SOURCE PPT from Azure Blob: {SOURCE_CONTAINER}/{blob_name}")
    except Exception as e:
        logger.exception(f"Failed to delete SOURCE PPT from Azure Blob: {blob_name}")
        raise e
