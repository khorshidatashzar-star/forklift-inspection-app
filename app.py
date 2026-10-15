                                         SET status='Completed',updated_at=?
                                         WHERE id=?""",(now(),r.id))
                            c.execute("""INSERT INTO asset_movements(
                                         serial_number,from_location,to_location,
                                         case_number,approved_by,changed_at)
                                         VALUES(?,?,?,?,?,?)""",
                                      (r.serial_number,old,r.requested_location,
                                       r.case_number,u["full_name"],now()))
                            c.commit();c.close()

                            audit(r.case_number,"IT",u["full_name"],
                                  "Asset Location Updated",
                                  f"{old} -> {r.requested_location}")
                            st.success("Asset location updated.")
                            st.rerun()

    with tab_overview:
        st.subheader("System Overview")

        users=query("SELECT * FROM users")
        eq=equipment()
        active_assets=eq[~eq.status.astype(str).str.lower().isin(["deleted"])]
        pending=query("""SELECT * FROM asset_review
                         WHERE status='Waiting IT Location Update'""")

        a,b,c,d=st.columns(4)
        a.metric("Active Staff",len(users[users.active.fillna(1).astype(int)==1]) if not users.empty else 0)
        b.metric("Active Machines",len(active_assets))
        c.metric("Active Locations",active_assets.location.dropna().nunique())
        d.metric("Open IT Requests",len(pending))

        st.markdown("### Assets by Location")
        if active_assets.empty:
            st.info("No active assets.")
        else:
            by_loc=(active_assets.groupby("location",dropna=False)
                    .size().reset_index(name="Machines")
                    .rename(columns={"location":"Location"}))
            st.dataframe(by_loc,hide_index=True,use_container_width=True)

        st.markdown("### Staff by Role")
        if users.empty:
            st.info("No staff users.")
        else:
            by_role=(users.groupby("role",dropna=False)
                     .size().reset_index(name="Staff")
                     .rename(columns={"role":"Role"}))
            st.dataframe(by_role,hide_index=True,use_container_width=True)


def normalize_user_access(u):
    """Enforce role/location scope regardless of what is stored in the user record."""
    if not u:
        return u

    user = dict(u)
    role = str(user.get("role") or "").strip()

    if role in ("Maintenance", "Compliance", "Compliance & Audit", "IT"):
        user["location"] = "ALL"

    return user


def reminder_status():
    c = conn()
    due = c.execute("""
        SELECT COUNT(*) AS n
        FROM inspections
        WHERE status='Repairing'
          AND next_reminder_at IS NOT NULL
          AND next_reminder_at<=?
    """, (now(),)).fetchone()["n"]
    c.close()
    return due

init_db()

st.markdown("""
<div class='hero'>
    <h2>🏗️ 18W Forklift Safety Portal</h2>
    <div class='muted'>
        Daily inspection • repair control • compliance • role-based access
    </div>
</div>
""", unsafe_allow_html=True)

u = normalize_user_access(st.session_state.get("user"))

if u:
    # Keep the normalized access scope in session as well.
    st.session_state["user"] = u
    logout_button()

    if force_password_change(u):
        st.stop()

    role = u["role"]

    if role == "WH Manager":
        manager_actions(u)

    elif role == "Maintenance":
        corrective_actions(u)

    elif role in ("Compliance", "Compliance & Audit"):
        compliance_actions(u)

    elif role == "IT":
        it_actions(u)

    else:
        st.error("Invalid role.")

else:
    mode = st.sidebar.radio(
        "Open",
        ["🚜 Start Inspection", "🔐 Staff Login"]
    )

    if mode == "🚜 Start Inspection":
        run_driver()
    else:
        temporary_role_test_access()
        st.divider()
        staff_login()
    
