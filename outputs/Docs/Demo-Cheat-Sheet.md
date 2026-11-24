# Admin Console -- demo cheat sheet

One page of framing, one page of run order, one page of answers. Glance, don't read.

---

## The one sentence

**A single console that puts every routine Microsoft 365 and Exchange task behind
a button, with the same safety rails and the same audit trail on every one.**

If you only get thirty seconds, that is the sentence.

## The three claims to prove

1. **It is faster.** Multi-step jobs that took a checklist now take a button.
2. **It is safer.** The tools refuse to do the wrong thing, and say why.
3. **It travels.** Same toolkit, six tenants, zero code changes -- config only.

Everything you demo should be visibly serving one of those three. If a screen
doesn't, skip it.

## Who it is for

Service desk and IT admins who need to do the task without needing to know the
cmdlet. Not a replacement for PowerShell -- a floor under it.

---

## Run order

Roughly 15 minutes. Cut from the bottom if time runs short.

### 1. Open cold (1 min)

Show the console before connecting. Point at the header: tenant name, connection
state. **Point out what you cannot press yet.** That is the first safety claim,
made before you've claimed anything.

> "It won't let you run a cloud tool with no cloud session. It's not guessing --
> it's asking the service."

### 2. Connect (2 min)

Activate (PIM) -> Connect ExO -> Connect SPO.

The console closes during sign-in and comes back. **Say that before it happens**,
or it looks like a crash:

> "It gets out of the way so the sign-in window can take focus. It comes back on
> its own."

### 3. Find something fast (1 min)

Type in the search box. Show it matching **by button name and by script name** --
type "not" and Returning User appears, because it calls NotDeadYet.ps1.

Then open the **Rooms** group. Five room tools behind one button.

> "Rooms matter, but they don't deserve five slots on the front page."

### 4. A read-only tool (2 min)

**Mailbox Health** or **Message Trace**. Pick one you can run against a real
object without hesitating.

Press the **?** while it's open.

> "Every tool has this. What it does, what it touches, and what will bite you.
> Written for the person who's never opened it before."

That single click is the strongest thing in the demo. It answers the question
every manager in the room is actually asking: *what happens when Glenn's on
holiday.*

### 5. A tool that refuses (2 min) -- THE MOMENT

**Manage Mail Object**, Remove, pointed at a real person's address with
**Shared mailbox** selected.

It looks the object up, tells you it found a UserMailbox, and **refuses**.

> "Typing the wrong address into a delete is how you lose someone's mailbox. It
> doesn't ask you to be careful -- it checks."

If you demo one write-capable tool, demo this one. Rehearse it; make sure the
address you use really is a user mailbox.

### 6. A real workflow (3 min)

**Termination** (Disable-User). Walk the steps without necessarily running it:
sign-in block, licence removal, forwarding, convert to shared, hide from GAL.

Two things to name:

- Licence groups are flagged **RemoveOnTerm** in config -- it removes the right
  one because someone told it which, not because it guessed.
- Anything that fails goes to a **retry queue** rather than vanishing.

### 7. The audit log (1 min)

Open the CSV or the Report Browser. Timestamp, operator, machine, **tenant**,
script, action, target, result.

> "Every write, every tool, one format. Including the failures -- especially the
> failures."

### 8. The Dashboard -- state and configuration on one screen (3 min)

Open **Dashboard**. Top half is what's true right now: connections, toolkit
integrity, today's activity, the retry queue. Bottom is **Configure** -- every
editor in one row: Tool Manager, Tenant Manager, Theme Manager, QuickLaunch,
Tool Status.

Open **Tenant Manager** from there. Capabilities, directories, licence groups.

> "Six tenants run this. One of them has no on-premises anything. The console
> doesn't grey those tools out -- it doesn't show them at all. If it doesn't
> apply, you never see it."

Then **Run Tester**, from the same screen.

> "It checks itself. That's how I know a copy someone else is running is intact."

This is the strongest single screen you have, and it answers the question
nobody asks out loud: what happens when I'm not here. State on top, the means
to change it underneath, and a self-check button. One window.

### 9. Tool Status (1 min) -- only if they ask "how big is this?"

From Configure. Every script, what loads it, when it last ran, and its own help
one click away.

> "Ninety scripts, and it can tell me which ones are actually wired in. Anything
> that isn't gets parked, not deleted."

---

## Do not demo

- **Search and Purge** -- parked, pending role approval. Say "in progress" if asked.
- **Change Meeting Organizer** -- unresolved service-side error.
- Anything you haven't run once this week. Live demos punish improvisation.

## If something breaks

Don't debug in front of people. One line, move on:

> "That's what the retry queue is for -- I'll pick it up after."

A tool that fails *visibly and safely* is on-message. You built it to fail loudly
rather than quietly. Say so and keep going.

If a connection drops: reconnect, and note that nothing was half-done, because
the failure was logged.

---

## Likely questions

**"What if you get hit by a bus?"**
Every tool has a help button, there's a setup guide and a quick start, and the
console self-tests. The knowledge is in the folder, not in my head.

**"Could the service desk use this?"**
That's who it's for. Rights are still governed by their roles -- the console
doesn't grant anything, it just stops them needing to know the syntax.

**"What stops someone doing something catastrophic?"**
Three layers: buttons hidden or disabled when they don't apply, tools that verify
the object before acting, and confirmations on anything irreversible. Plus the
audit log -- everything is attributable.

**"Is this supported? What if it breaks?"**
It's PowerShell against documented Microsoft cmdlets. Nothing exotic, nothing
undocumented. Worst case it stops working and you do it the old way.

**"How long did this take?"**
Answer honestly and immediately follow with what it saves. The interesting number
is per-termination, per-onboarding, per-ticket -- not hours spent building.

**"Why not just buy something?"**
Fair question -- give it a straight answer. Commercial tools cover reporting well.
This covers *our* workflows, across *our* six tenants, with our termination steps
and our OU layout. And it's already deployed.

**"Does it work for the other tenants?"**
Yes -- that's the config split. Same scripts, different Tenant.json.

---

## Closing line

> "Nothing here is clever. It's the same tasks we already do, with the
> checklist built in and a record of who did what."

Understating it lands better than overselling it, and it's true.

---

## Before you start

- [ ] Publish-Toolkit run, files unblocked
- [ ] Config backed up
- [ ] Test account and test room picked, and **tried once today**
- [ ] The refuse-to-delete address confirmed to be a user mailbox
- [ ] PIM activated, or you'll demo the greyed-out state longer than planned
- [ ] Second monitor mirrored, not extended
