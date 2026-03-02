// threaded_comment - A module for creating Excel ThreadedComments.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#![warn(missing_docs)]

mod tests;

use std::io::Cursor;

use crate::xmlwriter::{
    xml_data_element, xml_declaration, xml_empty_tag, xml_end_tag, xml_start_tag,
};

/// The identity provider type for a threaded comment person.
///
/// This enum defines the identity provider used to authenticate the person
/// who created or is mentioned in a threaded comment.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum IdentityProvider {
    /// No identity provider - userId is the display name.
    #[default]
    None,

    /// Active Directory - userId is the SID.
    ActiveDirectory,

    /// Windows Live ID - userId is a 64-bit signed decimal CID.
    WindowsLiveId,

    /// Office 365 - userId is `tenant::objectId::userId` format.
    Office365,

    /// People Picker - userId is an email address.
    PeoplePicker,
}

impl IdentityProvider {
    /// Get the provider ID string for XML serialization.
    pub(crate) fn to_provider_id(&self) -> &'static str {
        match self {
            IdentityProvider::None => "None",
            IdentityProvider::ActiveDirectory => "AD",
            IdentityProvider::WindowsLiveId => "Windows Live",
            IdentityProvider::Office365 => "AD",
            IdentityProvider::PeoplePicker => "PeoplePicker",
        }
    }
}

/// Represents a person (author or mentioned user) in threaded comments.
///
/// Each person in the workbook's threaded comments has a unique ID and
/// can be referenced by comments and mentions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThreadedCommentPerson {
    /// Unique identifier (UUID format with braces).
    pub(crate) id: String,

    /// Display name shown in the comment.
    pub(crate) display_name: String,

    /// Provider-specific user ID.
    pub(crate) user_id: String,

    /// Identity provider type.
    pub(crate) provider_id: IdentityProvider,
}

impl ThreadedCommentPerson {
    /// Create a new person with the given display name.
    ///
    /// # Arguments
    ///
    /// * `display_name` - The name shown in comments (e.g., "John Smith").
    ///
    /// # Examples
    ///
    /// ```
    /// use rust_xlsxwriter::ThreadedCommentPerson;
    ///
    /// let person = ThreadedCommentPerson::new("John Smith");
    /// ```
    pub fn new(display_name: impl Into<String>) -> ThreadedCommentPerson {
        let display_name = display_name.into();
        ThreadedCommentPerson {
            id: Self::generate_guid(),
            display_name: display_name.clone(),
            user_id: display_name,
            provider_id: IdentityProvider::None,
        }
    }

    /// Set the user ID for this person.
    ///
    /// The format depends on the identity provider:
    /// - `None`: Same as display name
    /// - `ActiveDirectory`: SID string
    /// - `WindowsLiveId`: 64-bit decimal CID
    /// - `Office365`: `tenant::objectId::userId` format
    /// - `PeoplePicker`: Email address
    ///
    /// # Arguments
    ///
    /// * `user_id` - The provider-specific user identifier.
    pub fn set_user_id(mut self, user_id: impl Into<String>) -> ThreadedCommentPerson {
        self.user_id = user_id.into();
        self
    }

    /// Set the identity provider for this person.
    ///
    /// # Arguments
    ///
    /// * `provider` - The identity provider type.
    pub fn set_provider(mut self, provider: IdentityProvider) -> ThreadedCommentPerson {
        self.provider_id = provider;
        self
    }

    /// Get the person's unique ID.
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Get the person's display name.
    pub fn display_name(&self) -> &str {
        &self.display_name
    }

    /// Generate a new GUID in the format `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`.
    fn generate_guid() -> String {
        use std::time::{SystemTime, UNIX_EPOCH};

        // Simple pseudo-random GUID generation based on time and counter
        let duration = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default();

        let nanos = duration.as_nanos();
        let secs = duration.as_secs();

        // Use a static counter for uniqueness within the same nanosecond
        use std::sync::atomic::{AtomicU64, Ordering};
        static COUNTER: AtomicU64 = AtomicU64::new(0);
        let count = COUNTER.fetch_add(1, Ordering::Relaxed);

        format!(
            "{{{:08X}-{:04X}-{:04X}-{:04X}-{:012X}}}",
            (nanos & 0xFFFFFFFF) as u32,
            ((secs >> 16) & 0xFFFF) as u16,
            ((secs & 0xFFFF) | 0x4000) as u16,        // Version 4
            ((count >> 48) & 0x3FFF | 0x8000) as u16, // Variant
            (count & 0xFFFFFFFFFFFF) as u64
        )
    }
}

/// Represents a mention (@user) within a threaded comment.
///
/// Mentions link to a person in the comment text using character positions.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ThreadedCommentMention {
    /// Unique identifier for this mention.
    pub(crate) mention_id: String,

    /// ID of the person being mentioned.
    pub(crate) mention_person_id: String,

    /// 0-based character position where the mention starts in the text.
    pub(crate) start_index: u32,

    /// Number of characters in the mention (including @ prefix).
    pub(crate) length: u32,
}

impl ThreadedCommentMention {
    /// Create a new mention.
    ///
    /// # Arguments
    ///
    /// * `person` - The person being mentioned.
    /// * `start_index` - Character position where the mention starts (0-based).
    /// * `length` - Number of characters in the mention span.
    pub fn new(
        person: &ThreadedCommentPerson,
        start_index: u32,
        length: u32,
    ) -> ThreadedCommentMention {
        ThreadedCommentMention {
            mention_id: ThreadedCommentPerson::generate_guid(),
            mention_person_id: person.id.clone(),
            start_index,
            length,
        }
    }
}

/// Represents a single comment within a threaded comment thread.
///
/// Comments can have text, mentions, and link to parent comments for threading.
#[derive(Debug, Clone)]
pub struct ThreadedComment {
    /// Unique identifier for this comment.
    pub(crate) id: String,

    /// Cell reference (e.g., "A1").
    pub(crate) cell_ref: String,

    /// ID of the comment author.
    pub(crate) person_id: String,

    /// Comment text content.
    pub(crate) text: String,

    /// Creation timestamp (ISO 8601 format).
    pub(crate) date_created: String,

    /// ID of parent comment (for threading).
    pub(crate) parent_id: Option<String>,

    /// Whether the thread is resolved.
    pub(crate) done: Option<bool>,

    /// Mentions within the comment text.
    pub(crate) mentions: Vec<ThreadedCommentMention>,
}

impl ThreadedComment {
    /// Create a new threaded comment.
    ///
    /// # Arguments
    ///
    /// * `cell_ref` - The cell address (e.g., "A1").
    /// * `author` - The person creating the comment.
    /// * `text` - The comment text.
    ///
    /// # Examples
    ///
    /// ```
    /// use rust_xlsxwriter::{ThreadedComment, ThreadedCommentPerson};
    ///
    /// let author = ThreadedCommentPerson::new("John Smith");
    /// let comment = ThreadedComment::new("A1", &author, "Please review this cell");
    /// ```
    pub fn new(
        cell_ref: impl Into<String>,
        author: &ThreadedCommentPerson,
        text: impl Into<String>,
    ) -> ThreadedComment {
        ThreadedComment {
            id: ThreadedCommentPerson::generate_guid(),
            cell_ref: cell_ref.into(),
            person_id: author.id.clone(),
            text: text.into(),
            date_created: Self::current_timestamp(),
            parent_id: None,
            done: None,
            mentions: Vec::new(),
        }
    }

    /// Set the parent comment ID for threading.
    ///
    /// This links this comment as a reply to another comment.
    pub fn set_parent(mut self, parent: &ThreadedComment) -> ThreadedComment {
        self.parent_id = Some(parent.id.clone());
        self
    }

    /// Mark the thread as resolved/done.
    pub fn set_resolved(mut self, resolved: bool) -> ThreadedComment {
        self.done = Some(resolved);
        self
    }

    /// Add a mention to this comment.
    ///
    /// # Arguments
    ///
    /// * `mention` - The mention to add.
    pub fn add_mention(mut self, mention: ThreadedCommentMention) -> ThreadedComment {
        self.mentions.push(mention);
        self
    }

    /// Get the comment's unique ID.
    pub fn id(&self) -> &str {
        &self.id
    }

    /// Generate current timestamp in ISO 8601 format.
    fn current_timestamp() -> String {
        use std::time::{SystemTime, UNIX_EPOCH};

        let duration = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap_or_default();

        let secs = duration.as_secs();
        let millis = duration.subsec_millis();

        // Convert to approximate date/time (simplified)
        let days = secs / 86400;
        let remaining = secs % 86400;
        let hours = remaining / 3600;
        let minutes = (remaining % 3600) / 60;
        let seconds = remaining % 60;

        // Approximate year/month/day from days since epoch
        let mut year = 1970u32;
        let mut remaining_days = days as u32;

        loop {
            let days_in_year = if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) {
                366
            } else {
                365
            };
            if remaining_days < days_in_year {
                break;
            }
            remaining_days -= days_in_year;
            year += 1;
        }

        let days_in_months: [u32; 12] = if year % 4 == 0 && (year % 100 != 0 || year % 400 == 0) {
            [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        } else {
            [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
        };

        let mut month = 1u32;
        for &days in &days_in_months {
            if remaining_days < days {
                break;
            }
            remaining_days -= days;
            month += 1;
        }
        let day = remaining_days + 1;

        format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}.{:02}",
            year,
            month,
            day,
            hours,
            minutes,
            seconds,
            millis / 10
        )
    }
}

/// A collection of threaded comments for a worksheet.
///
/// This struct manages the XML serialization of threaded comments.
pub struct ThreadedComments {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) comments: Vec<ThreadedComment>,
    pub(crate) persons: Vec<ThreadedCommentPerson>,
}

impl ThreadedComments {
    /// Create a new ThreadedComments collection.
    pub(crate) fn new() -> ThreadedComments {
        ThreadedComments {
            writer: Cursor::new(Vec::with_capacity(4096)),
            comments: Vec::new(),
            persons: Vec::new(),
        }
    }

    /// Add a person to the collection if not already present.
    pub(crate) fn add_person(&mut self, person: ThreadedCommentPerson) {
        if !self.persons.iter().any(|p| p.id == person.id) {
            self.persons.push(person);
        }
    }

    /// Add a comment to the collection.
    pub(crate) fn add_comment(&mut self, comment: ThreadedComment) {
        self.comments.push(comment);
    }

    /// Check if there are any threaded comments.
    pub(crate) fn is_empty(&self) -> bool {
        self.comments.is_empty()
    }

    /// Assemble the XML file for threaded comments.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_threaded_comments();

        xml_end_tag(&mut self.writer, "ThreadedComments");
    }

    /// Write the <ThreadedComments> root element.
    fn write_threaded_comments(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
            ),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
        ];

        xml_start_tag(&mut self.writer, "ThreadedComments", &attributes);

        // Write each comment
        for comment in self.comments.clone() {
            self.write_threaded_comment(&comment);
        }
    }

    /// Write a single <threadedComment> element.
    fn write_threaded_comment(&mut self, comment: &ThreadedComment) {
        let mut attributes: Vec<(&str, String)> = vec![
            ("ref", comment.cell_ref.clone()),
            ("dT", comment.date_created.clone()),
            ("personId", comment.person_id.clone()),
            ("id", comment.id.clone()),
        ];

        if let Some(ref parent_id) = comment.parent_id {
            attributes.push(("parentId", parent_id.clone()));
        }

        if let Some(done) = comment.done {
            attributes.push((
                "done",
                if done {
                    "1".to_string()
                } else {
                    "0".to_string()
                },
            ));
        }

        let attr_refs: Vec<(&str, &str)> =
            attributes.iter().map(|(k, v)| (*k, v.as_str())).collect();

        xml_start_tag(&mut self.writer, "threadedComment", &attr_refs);

        // Write text element
        xml_data_element(
            &mut self.writer,
            "text",
            &comment.text,
            &[] as &[(&str, &str)],
        );

        // Write mentions if any
        if !comment.mentions.is_empty() {
            xml_start_tag(&mut self.writer, "mentions", &[] as &[(&str, &str)]);
            for mention in &comment.mentions {
                self.write_mention(mention);
            }
            xml_end_tag(&mut self.writer, "mentions");
        }

        xml_end_tag(&mut self.writer, "threadedComment");
    }

    /// Write a <mention> element.
    fn write_mention(&mut self, mention: &ThreadedCommentMention) {
        let attributes = [
            ("mentionpersonId", mention.mention_person_id.as_str()),
            ("mentionId", mention.mention_id.as_str()),
            ("startIndex", &mention.start_index.to_string()),
            ("length", &mention.length.to_string()),
        ];

        xml_empty_tag(&mut self.writer, "mention", &attributes);
    }
}

/// A collection of persons for the workbook.
///
/// This is stored separately from threaded comments at the workbook level.
pub struct ThreadedCommentPersons {
    pub(crate) writer: Cursor<Vec<u8>>,
    pub(crate) persons: Vec<ThreadedCommentPerson>,
}

impl ThreadedCommentPersons {
    /// Create a new person collection.
    pub(crate) fn new() -> ThreadedCommentPersons {
        ThreadedCommentPersons {
            writer: Cursor::new(Vec::with_capacity(2048)),
            persons: Vec::new(),
        }
    }

    /// Add a person to the collection.
    pub(crate) fn add_person(&mut self, person: ThreadedCommentPerson) {
        if !self.persons.iter().any(|p| p.id == person.id) {
            self.persons.push(person);
        }
    }

    /// Check if there are any persons.
    pub(crate) fn is_empty(&self) -> bool {
        self.persons.is_empty()
    }

    /// Assemble the XML file for persons.
    pub(crate) fn assemble_xml_file(&mut self) {
        xml_declaration(&mut self.writer);

        self.write_person_list();

        xml_end_tag(&mut self.writer, "personList");
    }

    /// Write the <personList> root element.
    fn write_person_list(&mut self) {
        let attributes = [
            (
                "xmlns",
                "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments",
            ),
            (
                "xmlns:x",
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            ),
        ];

        xml_start_tag(&mut self.writer, "personList", &attributes);

        for person in self.persons.clone() {
            self.write_person(&person);
        }
    }

    /// Write a <person> element.
    fn write_person(&mut self, person: &ThreadedCommentPerson) {
        let provider_id = person.provider_id.to_provider_id();

        let attributes = [
            ("displayName", person.display_name.as_str()),
            ("id", person.id.as_str()),
            ("userId", person.user_id.as_str()),
            ("providerId", provider_id),
        ];

        xml_empty_tag(&mut self.writer, "person", &attributes);
    }
}

impl Clone for ThreadedComments {
    fn clone(&self) -> Self {
        ThreadedComments {
            writer: Cursor::new(self.writer.get_ref().clone()),
            comments: self.comments.clone(),
            persons: self.persons.clone(),
        }
    }
}

impl Clone for ThreadedCommentPersons {
    fn clone(&self) -> Self {
        ThreadedCommentPersons {
            writer: Cursor::new(self.writer.get_ref().clone()),
            persons: self.persons.clone(),
        }
    }
}
