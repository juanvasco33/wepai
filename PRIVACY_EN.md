# WEP AI — Privacy Policy

**Effective date:** May 21, 2026
**Last updated:** May 21, 2026

This Privacy Policy explains what information WEP AI ("we", "us") collects about you, how we use it, and your rights regarding that information. It applies to your use of the WEP AI desktop application and related services ("the Service").

## Summary in plain English

- Most of your data stays on your computer. We do not run a cloud database for your documents.
- The prompts you send to the AI are sent to **Anthropic** (the company behind Claude) for processing. We do not store them on our servers.
- If you subscribe to a paid plan, **Stripe** processes your payment. We never see your full card number.
- We do not sell your personal information. We do not show ads.
- You can delete your account at any time.

## 1. What we collect

### Information you provide
- **Account information**: name, last name, email, password (stored as bcrypt hash, never plain text)
- **Plan selection**: free, Pro or Enterprise
- **Conversation prompts**: the natural-language requests you send to generate documents

### Information collected automatically
- **Usage metrics**: number of documents generated per month (for plan enforcement only)
- **Application logs**: timestamps of generations, errors and bug detector results — stored locally

### Information from payments (only if you subscribe)
- **Last 4 digits of your card**: stored locally to display in your billing UI ("Visa ending in 1234")
- **Stripe customer ID and subscription ID**: stored locally so the app can fetch your plan status

We do **not** store your full card number, CVV, expiry date, or billing address. These are handled exclusively by Stripe.

## 2. Where your data lives

| Data | Where it lives | Who can access |
|---|---|---|
| Account email + password hash | Local SQLite at `~/.wepai/wepai.db` | You |
| Conversation history | Local SQLite at `~/.wepai/wepai.db` | You |
| Generated documents | Local folder `~/Documents/WEP_AI/` | You |
| Payment data | Stripe's PCI-compliant systems | Stripe + you |
| Prompts (transient) | Sent to Anthropic's API at the moment of generation | Anthropic |

We do **not** have a cloud database with your data. If you uninstall WEP AI and delete `~/.wepai/` and `~/Documents/WEP_AI/`, you have effectively erased all your data from our side.

## 3. How prompts reach Anthropic

When you submit a prompt (e.g. "Generate an invoice for client ABC for $5,000"), the WEP AI app sends that prompt directly from your computer to Anthropic's Claude API. Anthropic processes the prompt and returns the structured JSON used to build your document.

Anthropic's handling of your prompt is governed by [Anthropic's Privacy Policy](https://www.anthropic.com/legal/privacy) and [Commercial Terms of Service](https://www.anthropic.com/legal/commercial-terms). Per Anthropic's policy, prompts sent through the API are not used to train Claude.

If you do not want your prompts to leave your computer, do not use WEP AI.

## 4. How we use your information

We use locally-stored information to:
- Authenticate you and enforce your plan limits
- Show you your conversation history and generated documents
- Detect and report bugs (the Excel validator runs locally)
- For paid plans: verify your subscription is active

We do **not** use your information to:
- Train AI models
- Build profiles for advertising
- Sell to third parties
- Share with brokers

## 5. Sharing of information

We share information only in these limited cases:

- **Anthropic** receives your prompts at the moment of generation, as described in section 3
- **Stripe** receives your payment information when you subscribe
- **Law enforcement** in response to a valid subpoena or court order
- **Successor company** if WEP AI is acquired or merged, your data may transfer under the same terms

We do not share with any other third party.

## 6. Your rights

Depending on where you live, you have the following rights:

### For all users
- **Access**: see the data we have about you (stored locally, so just open `~/.wepai/wepai.db` with any SQLite browser)
- **Deletion**: delete your account from the app, or manually delete `~/.wepai/` and `~/Documents/WEP_AI/`
- **Export**: copy your documents from `~/Documents/WEP_AI/` and your DB; both use open formats

### For California residents (CCPA)
You have the right to:
- Know what personal information we collect (see section 1)
- Know whether we sell or share it (we do not)
- Delete your personal information (see Deletion above)
- Opt out of sale of personal information (not applicable — we don't sell)
- Non-discrimination for exercising these rights

To exercise CCPA rights, email **privacy@wepai.app** from the email associated with your account.

### For European Union residents (GDPR)
You have the right to:
- Access your personal data (Article 15)
- Correct inaccurate data (Article 16)
- Erase your data ("right to be forgotten" — Article 17)
- Restrict processing (Article 18)
- Data portability (Article 20)
- Object to processing (Article 21)
- Withdraw consent at any time
- Lodge a complaint with your local supervisory authority

To exercise GDPR rights, email **privacy@wepai.app**. We respond within 30 days.

The legal basis for processing your data is: **contract performance** (you signed up for the Service) for account data; **legitimate interest** (improving the product) for anonymous usage stats; **consent** for optional features.

## 7. Children's privacy

WEP AI is not intended for use by children under 18. We do not knowingly collect data from children. If we discover an account belongs to a child, we will delete it.

## 8. Data retention

- **Free plan**: conversation history kept for 7 days (then auto-deleted locally)
- **Pro plan**: conversation history kept for 90 days
- **Enterprise plan**: conversation history kept for 1 year
- **Generated documents**: kept indefinitely (they live in your `~/Documents` folder; we never delete them)
- **Payment records**: kept by Stripe per their retention policy; we keep only your subscription ID and last 4 digits of card

When you delete your account, we delete locally-stored data immediately. Stripe retains payment records as required by financial regulations (typically 7 years).

## 9. Security

- Passwords are hashed with **bcrypt** (industry standard, not reversible)
- The local SQLite database has filesystem permissions limited to your user account
- Payment data is never stored in our systems; Stripe handles all card data on PCI-compliant infrastructure
- The Stripe webhook server (if deployed) verifies signatures on all incoming events to prevent forgery

Despite these measures, no system is 100% secure. If we become aware of a breach affecting your data, we will notify you within 72 hours as required by GDPR.

## 10. International transfers

If you use WEP AI from outside the United States, your prompts are sent to Anthropic's servers, which may be located in the US. We rely on Anthropic's safeguards for international transfers (see Anthropic's privacy policy).

For European users, this transfer is covered under Standard Contractual Clauses approved by the European Commission.

## 11. Changes to this Policy

We may update this Privacy Policy. Material changes will be announced in the app and via email at least 14 days before they take effect. Your continued use after changes constitutes acceptance.

## 12. Contact us

For any privacy-related questions:
- **Email**: privacy@wepai.app
- **Mailing address**: [WEP AI · address · city, state, ZIP, country]
- **Data Protection Officer** (for EU/GDPR matters): dpo@wepai.app
