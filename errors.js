// Log any errors to Error Sheet
const logError = (error) => {
    errorsSheet.appendRow([error]);
};

const errorHandler = {
    logAPIError: logError
};