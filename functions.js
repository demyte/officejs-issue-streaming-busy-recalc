// Immediately publish the supplied code to a streaming invocation.
function member(hierarchyOrMember, memberCode, relation, ordinal, invocation) {
  invocation.onCanceled = () => {};
  invocation.setResult([[String(memberCode?.[0]?.[0] ?? "")]]);
}

Office.onReady(() => CustomFunctions.associate("MEMBER", member));
