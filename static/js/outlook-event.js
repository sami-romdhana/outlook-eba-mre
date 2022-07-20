function addSignature(eventObj) {
  Office.context.mailbox.item.body.setSignatureAsync(
    "Hello World! This is a signature.",
    {
      coercionType: "html",
    },
    function () {
      eventObj.completed();
    }
  );
}

Office.actions.associate("addSignature", addSignature);
