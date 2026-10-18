// Kept separate so normal foreground session checks do not load case/identity logic.
export async function requestSummary(db,email){
 const row=await db.prepare('SELECT version,unread FROM support_inboxes WHERE email=?').bind(email).first();
 return {version:Number(row?.version||0),unread:Number(row?.unread||0)};
}
