def pick_insertion_date(case_block, queue_date):
    """
    Picks a realistic insertion date for a new note:
    - If queue_date exists, insert a few days before it (but after the last note).
    - If queue_date missing, insert a few days after the last existing note.
    - If no valid dates at all, use today's date.
    """
    today = datetime.today()

    # Extract sorted note dates
    note_dates = case_block["Note Date "].dropna().sort_values()
    if note_dates.empty:
        return today

    last_note_date = note_dates.iloc[-1]

    if pd.notna(queue_date):
        # Insert between last note and queue_date (if possible)
        if last_note_date < queue_date:
            delta_days = (queue_date - last_note_date).days
            if delta_days > 3:
                # Random offset within that range
                offset = random.randint(1, min(10, delta_days - 1))
                return last_note_date + timedelta(days=offset)
            else:
                # Too close, just use 1 day before queue_date
                return queue_date - timedelta(days=1)
        else:
            # Last note is after queue_date (data anomaly)
            return last_note_date + timedelta(days=1)
    else:
        # No queue date, insert a few days after last note
        offset = random.randint(2, 7)
        return last_note_date + timedelta(days=offset)
