const handleDraftDocument = (data, row, columns) => {
    Logger.log("9. In draft handler")
    handleStatus.setStatus(row, columns, "Draft", "document.draft");
};

const draft = {
    draftStatus: handleDraftDocument
}