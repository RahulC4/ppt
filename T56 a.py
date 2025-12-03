def delete_ppt_from_chroma(ppt_name: str) -> None:
    """
    Delete all Chroma slide indexes that belong to the given PPT.
    Matches your metadata:
    metadata = { "ppt_name": blob_name, ... }
    """

    logger.info(f"Deleting Chroma indexes for PPT: {ppt_name}")

    try:
        # ✅ DIRECT DELETE — no pre-query (avoids Chroma API bug)
        collection.delete(where={"ppt_name": ppt_name})

        # ✅ IMPORTANT: Persist the deletion to disk
        chroma_client.persist()

        logger.info(f"✅ Successfully deleted Chroma indexes for PPT: {ppt_name}")

    except Exception as e:
        logger.exception(f"❌ Failed to delete Chroma indexes for PPT: {ppt_name}")
        raise e
