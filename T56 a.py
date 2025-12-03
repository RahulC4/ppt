def delete_ppt_from_chroma(ppt_name: str) -> None:
    """
    Delete all Chroma slide indexes that belong to the given PPT.

    This matches your ingestion metadata:
    metadata = {
        "ppt_name": blob_name,
        "slide_index": ...,
        "slide_id": ...,
        ...
    }
    """

    logger.info(f"Deleting Chroma indexes for PPT: {ppt_name}")

    try:
        # Safety check: see how many records exist before delete
        res = collection.query(where={"ppt_name": ppt_name}, n_results=1)
        existing_ids = res.get("ids", [[]])[0]

        if not existing_ids:
            logger.warning(f"No Chroma records found for PPT: {ppt_name}")
            return

        # ✅ ACTUAL DELETE
        collection.delete(where={"ppt_name": ppt_name})

        logger.info(f"✅ Successfully deleted Chroma indexes for PPT: {ppt_name}")

    except Exception as e:
        logger.exception(f"❌ Failed to delete Chroma indexes for PPT: {ppt_name}")
        raise e
