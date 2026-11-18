# 6. Relations between PocketBase, SharePoint and tenant users

PocketBase relation fields can only point at PocketBase collections. Everything outside PocketBase is therefore connected either by **stored identifiers** (resolved live through Graph) or by **mirroring** it into a collection first.

```text
 Entra ID user ──oid──▶ users.entraOid            (PocketBase account, created at first sign-in)
       │                      ▲
       │ mirror               │ account (relation, set when the person signs in)
       ▼                      │
 directory_users ─────────────┘
       ▲
       │ owner (relation, matched by email)
       │
 sp_items ◀── mirror ── SharePoint list item ── OwnerLookupId ──▶ User Information List ──▶ email
       ▲
       │ item (relation)
 item_notes ── author ──▶ users                     (data that only lives in PocketBase)
```

## 1. Entra user ↔ PocketBase account

- **Key:** Entra object ID `oid` → `users.entraOid` (unique index).
- **Set by:** `/api/entra-login` (from the token's `oid` or `/me`) and the built-in OAuth2 hook (`oauth2User.id`). Both paths find the same record.
- Use `oid`, not email or UPN: UPNs can be renamed and some accounts have no `mail`.

## 2. Tenant directory ↔ PocketBase

| Pattern | How | Use when |
| --- | --- | --- |
| Live | `GET /api/directory/users?search=` (Graph `/users`) | pickers, lookups of people who are not mirrored |
| Mirror | `directory_users`, refreshed by cron | you need relations, filters, `expand`, offline queries |

`directory_users.account → users` links a directory entry to the person's PocketBase account. It stays empty until they sign in, so **people who never use the app can still be related to data**.

```js
pocketBase.collection("directory_users").getList(1, 25, { expand: "account" })
```

## 3. SharePoint item ↔ PocketBase

| Pattern | Store in PocketBase | Pros / cons |
| --- | --- | --- |
| **Reference** | `siteId`, `listId`, `spItemId` as text | always current; no relation, no `expand`, each read calls Graph |
| **Mirror** (implemented) | `sp_items` record per list item, `fields` as JSON | real relations, filters and API rules; data is as fresh as the last sync |

Mirror records are unique per `(listId, spItemId)`, so several lists can share `sp_items`. For several lists with different schemas, one collection per list with typed fields is clearer.

## 4. SharePoint Person column ↔ tenant user

This is the non-obvious one. Graph returns a Person column only as a lookup ID:

```json
{ "Title": "Item 1", "OwnerLookupId": "6", "AuthorLookupId": "10", "EditorLookupId": "10" }
```

`6` is the item ID in the site's hidden **User Information List**. Resolution in [`lib/sharepoint.js`](../pocketbase/pb_hooks/lib/sharepoint.js):

1. Read the list's columns (`$expand=columns`); Person columns have `personOrGroup`. Skip `readOnly` ones (the built-in `Author`/`Editor`, "Created By"/"Modified By").
2. Find the User Information List by its internal name `users` (its title is localized, for example "Benutzerinformationsliste" in German):
   ```http
   GET /sites/{site-id}/lists?$select=id,name,system
   ```
3. Read its items: `fields(EMail, UserName, Title)`; map lookup ID → email.
4. Match the email to `directory_users.mail` **or** `directory_users.upn` (they differ for some accounts, and guests have UPNs like `name_domain#EXT#@tenant.onmicrosoft.com`).
5. Store `sp_items.ownerEmail` and the relation `sp_items.owner → directory_users`.

Result, from the test list:

| Item | Owner column | Mirror result |
| --- | --- | --- |
| owned by a user who signed in | set | `owner` → directory entry → `account` → PocketBase user |
| owned by a user who never signed in | set | `owner` → directory entry, `account` empty |
| no owner | empty | `owner` empty |

Queries this enables:

```js
// items with their owner and the owner's app account
pocketBase.collection("sp_items").getFullList({ expand: "owner,owner.account" })
// "my items"
pocketBase.collection("sp_items").getFullList({ filter: pocketBase.filter("owner.account = {:id}", { id: myUserId }) })
```

API rules can use the same path, for example a view rule that only shows your own items: `owner.account = @request.auth.id`.

**Writing a Person column:** set `<Column>LookupId` to the user's ID in the User Information List. A user who has never been referenced on the site has no entry there yet. Graph has no "ensure user" call; that needs the SharePoint REST API (`/_api/web/ensureuser`). Not tested here.

## 5. PocketBase data ↔ SharePoint item

Data that only belongs to the app (comments, approvals, ratings, local status) stays in PocketBase and points at the mirrored item:

- `item_notes.item → sp_items` (`cascadeDelete`: notes disappear if the mirror record is deleted)
- `item_notes.author → users`
- create rule `@request.body.author = @request.auth.id`: you can only write notes as yourself

```js
pocketBase.collection("sp_items").getOne(id, { expand: "item_notes_via_item" })   // back-relation
```

Caveat: the mirror never deletes `sp_items` records whose list item was deleted in SharePoint. Add that to the sync (compare IDs, or use delta queries) if notes should follow deletions.

## 6. SharePoint list ↔ SharePoint list (Lookup columns)

A Lookup column to another list appears the same way, as `<Column>LookupId` (the other list's item ID). Mirror both lists and resolve the ID to the other mirror record, like the Person column. Not tested here.

## Choosing

| Need | Use |
| --- | --- |
| Show current SharePoint data, no relations | live route |
| Filter/sort/expand across users and items, API rules | mirror |
| App-only data about SharePoint items | PocketBase collection with relation to the mirror |
| Relate people who never sign in | `directory_users`, not `users` |
