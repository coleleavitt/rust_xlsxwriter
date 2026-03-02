// Threaded comment unit tests.
//
// SPDX-License-Identifier: MIT OR Apache-2.0
//
// Copyright 2022-2026, John McNamara, jmcnamara@cpan.org

#[cfg(test)]
mod threaded_comment_tests {
    use crate::threaded_comment::{
        IdentityProvider, ThreadedComment, ThreadedCommentMention, ThreadedCommentPerson,
        ThreadedCommentPersons, ThreadedComments,
    };

    #[test]
    fn test_person_creation() {
        let person = ThreadedCommentPerson::new("John Smith");
        assert_eq!(person.display_name(), "John Smith");
        assert!(!person.id().is_empty());
        assert!(person.id().starts_with('{'));
        assert!(person.id().ends_with('}'));
    }

    #[test]
    fn test_person_with_provider() {
        let person = ThreadedCommentPerson::new("Jane Doe")
            .set_user_id("jane@company.com")
            .set_provider(IdentityProvider::PeoplePicker);

        assert_eq!(person.display_name(), "Jane Doe");
        assert_eq!(person.user_id, "jane@company.com");
        assert_eq!(person.provider_id, IdentityProvider::PeoplePicker);
    }

    #[test]
    fn test_comment_creation() {
        let author = ThreadedCommentPerson::new("Test Author");
        let comment = ThreadedComment::new("A1", &author, "Test comment");

        assert_eq!(comment.cell_ref, "A1");
        assert_eq!(comment.text, "Test comment");
        assert_eq!(comment.person_id, author.id);
        assert!(comment.parent_id.is_none());
        assert!(comment.done.is_none());
    }

    #[test]
    fn test_comment_threading() {
        let author = ThreadedCommentPerson::new("Author");
        let root_comment = ThreadedComment::new("B2", &author, "Root comment");
        let reply = ThreadedComment::new("B2", &author, "Reply").set_parent(&root_comment);

        assert!(reply.parent_id.is_some());
        assert_eq!(reply.parent_id.unwrap(), root_comment.id);
    }

    #[test]
    fn test_comment_resolved() {
        let author = ThreadedCommentPerson::new("Author");
        let comment = ThreadedComment::new("C3", &author, "Comment").set_resolved(true);

        assert_eq!(comment.done, Some(true));
    }

    #[test]
    fn test_mention_creation() {
        let person = ThreadedCommentPerson::new("Mentioned User");
        let mention = ThreadedCommentMention::new(&person, 0, 14);

        assert_eq!(mention.mention_person_id, person.id);
        assert_eq!(mention.start_index, 0);
        assert_eq!(mention.length, 14);
    }

    #[test]
    fn test_comment_with_mention() {
        let author = ThreadedCommentPerson::new("Author");
        let mentioned = ThreadedCommentPerson::new("John");
        let mention = ThreadedCommentMention::new(&mentioned, 0, 5);

        let comment =
            ThreadedComment::new("D4", &author, "@John please review").add_mention(mention);

        assert_eq!(comment.mentions.len(), 1);
        assert_eq!(comment.mentions[0].start_index, 0);
        assert_eq!(comment.mentions[0].length, 5);
    }

    #[test]
    fn test_threaded_comments_collection() {
        let mut tc = ThreadedComments::new();
        let author = ThreadedCommentPerson::new("Author");

        tc.add_person(author.clone());
        tc.add_comment(ThreadedComment::new("E5", &author, "Comment 1"));
        tc.add_comment(ThreadedComment::new("E6", &author, "Comment 2"));

        assert_eq!(tc.persons.len(), 1);
        assert_eq!(tc.comments.len(), 2);
        assert!(!tc.is_empty());
    }

    #[test]
    fn test_persons_collection() {
        let mut persons = ThreadedCommentPersons::new();

        let person1 = ThreadedCommentPerson::new("Person 1");
        let person2 = ThreadedCommentPerson::new("Person 2");

        persons.add_person(person1.clone());
        persons.add_person(person2);
        persons.add_person(person1); // Duplicate should be ignored

        assert_eq!(persons.persons.len(), 2);
    }

    #[test]
    fn test_identity_provider_to_string() {
        assert_eq!(IdentityProvider::None.to_provider_id(), "None");
        assert_eq!(IdentityProvider::ActiveDirectory.to_provider_id(), "AD");
        assert_eq!(
            IdentityProvider::WindowsLiveId.to_provider_id(),
            "Windows Live"
        );
        assert_eq!(IdentityProvider::Office365.to_provider_id(), "AD");
        assert_eq!(
            IdentityProvider::PeoplePicker.to_provider_id(),
            "PeoplePicker"
        );
    }
}
